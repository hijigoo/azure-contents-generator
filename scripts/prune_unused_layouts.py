"""실제로 사용되지 않는 slideLayout/slideMaster 와 그것들만 참조하는 미디어를 제거한다.

clean.py 가 layout 까지는 손대지 않아 layout 전용 큰 이미지(템플릿 배경 등)가
남는 경우가 많다. 본 스크립트는:
1. 모든 슬라이드의 .rels 에서 사용되는 layout 집합 계산
2. 사용되지 않는 layout 의 .xml / .rels / Content_Types / presentation rels 제거
3. 남은 .rels 들을 다시 스캔해 unreferenced 미디어 제거

단, 1개 master 는 보존 (PowerPoint 는 최소 1 master 필요).
"""
from __future__ import annotations
import argparse, os, re, shutil, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}


def read_all(z: zipfile.ZipFile) -> dict[str, bytes]:
    return {n: z.read(n) for n in z.namelist()}


def referenced_in(data: dict[str, bytes], suffix_glob: str) -> set[str]:
    pat = re.compile(rf'([^"/\s>]+\.{suffix_glob})')
    used: set[str] = set()
    for content in data.values():
        for m in pat.findall(content.decode("utf-8", errors="ignore")):
            used.add(m)
    return used


def basename(p: str) -> str:
    return os.path.basename(p)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True)
    args = ap.parse_args()
    src = Path(args.input)
    tmp = src.with_suffix(src.suffix + ".tmp")

    with zipfile.ZipFile(src) as z:
        data = read_all(z)

    # 1) 사용되는 slideLayout 수집 (slides → layout 참조)
    slide_rels = {n: c for n, c in data.items() if re.match(r"ppt/slides/_rels/slide\d+\.xml\.rels$", n)}
    used_layouts: set[str] = set()
    for content in slide_rels.values():
        for m in re.findall(r'slideLayouts/(slideLayout\d+\.xml)', content.decode("utf-8", "ignore")):
            used_layouts.add(m)

    all_layouts = sorted({basename(n) for n in data if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)})
    unused_layouts = [n for n in all_layouts if n not in used_layouts]
    print(f"[deadcode] 전체 layout {len(all_layouts)} / 사용중 {len(used_layouts)} / 미사용 {len(unused_layouts)}")

    # 2) 미사용 layout 의 .xml + .rels 제거
    removed_paths: set[str] = set()
    for layout in unused_layouts:
        removed_paths.add(f"ppt/slideLayouts/{layout}")
        removed_paths.add(f"ppt/slideLayouts/_rels/{layout}.rels")

    # 3) [Content_Types].xml 정리
    ct_path = "[Content_Types].xml"
    ct_xml = data[ct_path].decode("utf-8")
    for layout in unused_layouts:
        ct_xml = re.sub(
            rf'\s*<Override\s+PartName="/ppt/slideLayouts/{re.escape(layout)}"[^/]*/>',
            "",
            ct_xml,
        )
    data[ct_path] = ct_xml.encode("utf-8")

    # 4) presentation.xml.rels 안의 미사용 layout 참조 제거 (대체로 master 가 참조)
    for master_rels_name in [n for n in data if re.match(r"ppt/slideMasters/_rels/slideMaster\d+\.xml\.rels$", n)]:
        txt = data[master_rels_name].decode("utf-8")
        new = re.sub(
            r'<Relationship[^/]*Target="\.\./slideLayouts/(' + "|".join(map(re.escape, unused_layouts)) + r')"[^/]*/>',
            "",
            txt,
        ) if unused_layouts else txt
        data[master_rels_name] = new.encode("utf-8")

    # slideMaster 의 sldLayoutIdLst 항목도 정리해야 PowerPoint 가 안 깨짐.
    # 안전을 위해 master xml 의 sldLayoutId 중 rels 에서 사라진 rId 도 제거.
    for master_path in [n for n in data if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)]:
        rels_path = master_path.replace("slideMaster", "_rels/slideMaster") + ".rels"
        if rels_path not in data:
            continue
        rels_xml = data[rels_path].decode("utf-8")
        valid_rids = set(re.findall(r'Id="([^"]+)"', rels_xml))
        master_xml = data[master_path].decode("utf-8")
        # sldLayoutId 요소 중 r:id 가 valid_rids 에 없는 것 삭제
        def keep(match: re.Match) -> str:
            rid = re.search(r'r:id="([^"]+)"', match.group(0))
            return match.group(0) if rid and rid.group(1) in valid_rids else ""
        master_xml = re.sub(r'<p:sldLayoutId[^/]*/>', keep, master_xml)
        data[master_path] = master_xml.encode("utf-8")

    # 5) 새 zip 작성 (1차)
    for p in list(removed_paths):
        data.pop(p, None)

    # 6) 이제 모든 남은 .rels 를 스캔해 미참조 미디어 제거
    media_files = {n for n in data if n.startswith("ppt/media/")}
    rels_blob = b"".join(c for n, c in data.items() if n.endswith(".rels"))
    rels_text = rels_blob.decode("utf-8", "ignore")
    keep_media = set()
    for m in re.findall(r'media/([^"\s<>]+)', rels_text):
        keep_media.add(f"ppt/media/{m}")
    unused_media = sorted(media_files - keep_media)
    removed_bytes = sum(len(data[m]) for m in unused_media)
    for m in unused_media:
        data.pop(m, None)
    print(f"[deadcode] 미사용 layout 제거: {len(unused_layouts)}")
    print(f"[deadcode] 미사용 미디어 제거: {len(unused_media)}개 / {removed_bytes/1024/1024:.1f} MB")

    # Content_Types 에서 사라진 미디어 Override/Default 도 정리 (Default 는 ext 단위라 그대로 둠)
    for media in unused_media:
        ct_xml = data[ct_path].decode("utf-8")
        ct_xml = re.sub(
            rf'\s*<Override\s+PartName="/{re.escape(media)}"[^/]*/>',
            "",
            ct_xml,
        )
        data[ct_path] = ct_xml.encode("utf-8")

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in data.items():
            z.writestr(name, content)

    shutil.move(str(tmp), str(src))
    print(f"[deadcode] 저장: {src} ({src.stat().st_size/1024/1024:.1f} MB)")


if __name__ == "__main__":
    main()
