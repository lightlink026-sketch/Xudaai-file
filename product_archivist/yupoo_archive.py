from __future__ import annotations

import json
import re
import shutil
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

from core import IMAGE_EXTS, VIDEO_EXTS, file_sha256, missing_product_fields, safe_stem
from yupoo_tool import read_product_detail


def _product_code(product_id: str) -> str:
    return "P-" + uuid.UUID(product_id).hex[:12].upper()


def _copy_tree_files(source: Path, destination: Path) -> list[dict[str, str]]:
    copied = []
    if not source.is_dir():
        return copied
    for path in source.rglob("*"):
        if not path.is_file() or path.is_symlink():
            continue
        relative = path.relative_to(source)
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, target)
        copied.append({"path": target.relative_to(destination.parent).as_posix(), "sha256": file_sha256(target)})
    return copied


def standardize_yupoo_downloads(download_root: Path, archive_root: Path, selected_folders: list[Path] | None = None) -> list[Path]:
    """Convert collector output directly into canonical product archive folders."""
    download_root = Path(download_root).resolve()
    archive_root = Path(archive_root).resolve()
    selected = {Path(path).resolve() for path in selected_folders or []}
    manifests = sorted(download_root.rglob("info.json"))
    albums: list[tuple[Path, dict[str, Any]]] = []
    for manifest_path in manifests:
        try:
            data = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        url = str(data.get("url") or "").strip()
        album_match = re.search(r"/albums/(\d+)", url)
        if album_match and (not selected or manifest_path.parent.resolve() in selected):
            albums.append((manifest_path.parent, data))
    if not albums:
        raise ValueError("又拍下载目录中没有可转换的商品相册。")

    archives: list[Path] = []
    archive_root.mkdir(parents=True, exist_ok=True)
    for source_folder, album in albums:
        url = str(album.get("url") or "").strip()
        album_id = re.search(r"/albums/(\d+)", url).group(1)
        supplier = str(album.get("account") or source_folder.parent.name or "未指定供应商").strip()
        name = str(album.get("title") or source_folder.name).strip() or "未命名产品"
        product_id = str(uuid.uuid5(uuid.NAMESPACE_URL, f"yupoo:{url}"))
        product_code = _product_code(product_id)

        previous_path = None
        previous: dict[str, Any] = {}
        for manifest in archive_root.rglob("产品档案.json"):
            try:
                candidate = json.loads(manifest.read_text(encoding="utf-8"))
            except (OSError, ValueError):
                continue
            if candidate.get("product_id") == product_id:
                previous_path = manifest.parent
                previous = candidate
                break

        details = read_product_detail(source_folder)
        values = {
            "name": name,
            "category": "",
            "brand": "",
            "model": "",
            "sku": "",
            "supplier_name": supplier,
            "supplier_sku": "",
            "description": str(album.get("description") or "").strip(),
            "colors": "",
            "sizes": str(details.get("尺码/size") or "").strip(),
            "material": "",
            "unit": "pcs",
            "currency": "CNY" if details.get("单价/price(RMB)") else "",
            "unit_price": str(details.get("单价/price(RMB)") or "").strip(),
            "cost_price": str(details.get("单价/price(RMB)") or "").strip(),
            "wholesale_price": "",
            "moq": str(details.get("最小起定量/MOQ") or "").strip(),
            "carton_spec": str(details.get("箱规/Carton dimensions") or "").strip(),
            "units_per_carton": "",
            "carton_quantity": "",
            "carton_volume_cbm": str(details.get("单箱体积/CBM") or "").strip(),
            "gross_weight_kg_per_carton": str(details.get("单箱毛重/Gross weight kg") or "").strip(),
            "net_weight_kg_per_carton": str(details.get("单箱净重/Net weight kg") or "").strip(),
            "lead_time": "",
            "notes": str(details.get("备注/remark") or "").strip(),
            "field_sources": {key: "Yupoo 商品详情" for key, value in details.items() if str(value or "").strip()},
        }
        # A prior standard archive is the authority for non-empty human-confirmed fields.
        for key in values:
            if str(previous.get(key) or "").strip():
                values[key] = previous[key]
        if previous.get("cost_price") or previous.get("unit_price"):
            values["cost_price"] = previous.get("cost_price") or previous.get("unit_price")
        values["unit_price"] = values.get("cost_price", "")
        values["name"] = str(previous.get("name") or values["name"])
        values["supplier_name"] = str(previous.get("supplier_name") or values["supplier_name"])
        values["field_sources"] = {**values.get("field_sources", {}), **(previous.get("field_sources") or {})}

        destination = previous_path or (archive_root / safe_stem(values["supplier_name"]) / f"{product_code}-{safe_stem(values['name'])}")
        destination.parent.mkdir(parents=True, exist_ok=True)
        temp = destination.with_name(destination.name + ".converting")
        if temp.exists():
            shutil.rmtree(temp)
        pictures = temp / "图片"
        gallery = pictures / "副图"
        videos_dir = temp / "视频"
        specs_dir = temp / "规格与包装"
        qualifications_dir = temp / "资质与说明"
        for folder in (gallery, videos_dir, specs_dir, qualifications_dir):
            folder.mkdir(parents=True, exist_ok=True)

        image_entries = []
        cover = Path(str(album.get("cover_path") or ""))
        candidates = []
        if cover.is_file():
            candidates.append((cover, True))
        for image in album.get("images") or []:
            path = Path(str(image))
            if path.is_file() and path.resolve() != cover.resolve():
                candidates.append((path, False))
        # If the metadata is incomplete, use the downloaded image files in place.
        if not candidates:
            candidates = [(path, path.name.startswith("ZT")) for path in sorted(source_folder.iterdir())
                          if path.is_file() and path.suffix.lower() in IMAGE_EXTS]
        expected_count = int(album.get("expected_image_count") or 0)
        saved_count = int(album.get("image_count") or len(album.get("images") or []))
        download_complete = expected_count > 0 and saved_count == expected_count and len(candidates) >= expected_count
        if previous_path and not download_complete:
            _copy_tree_files(previous_path / "图片", pictures)
            image_entries = list(previous.get("images") or [])
        else:
            for index, (path, is_cover) in enumerate(candidates, 1):
                target = pictures / f"主图{path.suffix.lower()}" if is_cover else gallery / f"{index:03d}{path.suffix.lower()}"
                if target.exists():
                    target = target.with_name(f"{target.stem}_{file_sha256(path)[:8]}{target.suffix}")
                shutil.copy2(path, target)
                image_entries.append({"path": target.relative_to(temp).as_posix(), "role": "main" if is_cover else "gallery", "sha256": file_sha256(target)})

        # Keep curated supplemental files already attached to the standard archive.
        if previous_path and previous_path.is_dir():
            for folder_name in ("视频", "规格与包装", "资质与说明"):
                _copy_tree_files(previous_path / folder_name, temp / folder_name)
        video_entries = []
        for path in videos_dir.rglob("*"):
            if path.is_file() and path.suffix.lower() in VIDEO_EXTS:
                video_entries.append({"path": path.relative_to(temp).as_posix(), "sha256": file_sha256(path)})

        product = {
            **previous,
            **values,
            "schema_version": 1,
            "product_id": product_id,
            "product_code": str(previous.get("product_code") or product_code),
            "name": values["name"],
            "supplier_name": values["supplier_name"],
            "source_type": "Yupoo",
            "source_url": url,
            "source_product_id": album_id,
            "updated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "review_status": "needs_info" if missing_product_fields(values) else "ready_for_review",
            "missing_fields": missing_product_fields(values),
            "collection_status": "complete" if download_complete else "partial",
            "images": image_entries,
            "videos": video_entries,
            "source_files": [],
        }
        (temp / "产品档案.json").write_text(json.dumps(product, ensure_ascii=False, indent=2), encoding="utf-8")
        if destination.exists():
            shutil.rmtree(destination)
        temp.rename(destination)
        archives.append(destination)

    return archives
