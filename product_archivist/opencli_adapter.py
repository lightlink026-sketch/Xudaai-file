from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any

from core import file_sha256, make_product_package, safe_stem, scan_product_folder


def _opencli(configured: str | None = None) -> str:
    if configured and configured.strip():
        value = configured.strip().strip('"')
        candidate = Path(value).expanduser()
        if candidate.exists() and candidate.is_file():
            return str(candidate.resolve())
        resolved = shutil.which(value)
        if resolved:
            return resolved
        raise RuntimeError(f"找不到配置的 OpenCLIApp：{value}")
    candidate = shutil.which("opencli") or shutil.which("opencli.cmd")
    if candidate:
        return candidate
    raise RuntimeError("没有找到 OpenCLI。请先安装 OpenCLIApp，并确保 opencli 命令可在 Windows 终端运行。")


def _run(args: list[str], timeout: int = 900, executable: str | None = None,
         profile: str | None = None, progress_file: Path | None = None,
         on_progress: Any | None = None) -> str:
    exe = _opencli(executable)
    command = [exe, *( ["--profile", profile] if profile else [] ), *args]
    if os.name == "nt" and Path(exe).suffix.lower() in {".cmd", ".bat"}:
        command = [os.environ.get("COMSPEC", "cmd.exe"), "/d", "/s", "/c", subprocess.list2cmdline(command)]
    progress_file = Path(progress_file) if progress_file else None
    last_progress = ""
    deadline = time.monotonic() + timeout
    started_at = time.time()
    with tempfile.TemporaryFile(mode="w+t", encoding="utf-8", errors="replace") as captured:
        process = subprocess.Popen(command, stdout=captured, stderr=subprocess.STDOUT, text=True)
        while process.poll() is None:
            current_progress = progress_file
            if progress_file and progress_file.is_dir():
                candidates = [path for path in progress_file.glob("*.progress.txt") if path.stat().st_mtime >= started_at - 1]
                current_progress = max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None
            if current_progress and on_progress and current_progress.is_file():
                try:
                    lines = current_progress.read_text(encoding="utf-8", errors="replace").splitlines()
                    current = lines[-1].strip() if lines else ""
                    if current and current != last_progress:
                        last_progress = current
                        on_progress(current)
                except OSError:
                    pass
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                process.kill()
                process.wait()
                raise subprocess.TimeoutExpired(command, timeout)
            try:
                process.wait(timeout=min(0.5, remaining))
            except subprocess.TimeoutExpired:
                continue
        captured.seek(0)
        output = captured.read().strip()
    if process.returncode:
        raise RuntimeError((output or f"OpenCLI 退出码 {process.returncode}").strip()[-5000:])
    return output


def check_opencli(executable: str | None = None) -> dict[str, str]:
    exe = _opencli(executable)
    try:
        output = _run(["--help"], timeout=20, executable=exe)
    except subprocess.TimeoutExpired:
        raise RuntimeError("OpenCLIApp 启动检查超时。")
    return {"path": exe, "output": output[:1000]}


def list_1688_profiles(executable: str | None = None) -> list[str]:
    """Return Browser Bridge profile ids connected to the local OpenCLI instance."""
    output = _run(["profile", "list"], timeout=20, executable=executable)
    return list(dict.fromkeys(re.findall(r"^\s*([A-Za-z0-9_-]+)\s+[—-]\s+connected\b", output, re.I | re.M)))


def resolve_1688_profile(executable: str | None = None, selected: str = "") -> str:
    """Choose the requested authenticated session, or the sole authenticated session."""
    profiles = list_1688_profiles(executable)
    if selected:
        if selected not in profiles:
            raise RuntimeError(f"浏览器会话 {selected} 已断开，请刷新会话列表后重试。")
        candidates = [selected]
    else:
        candidates = profiles
    authenticated: list[str] = []
    for profile in candidates:
        try:
            _run(["1688", "whoami", "--format", "json"], timeout=25, executable=executable, profile=profile)
            authenticated.append(profile)
        except RuntimeError:
            continue
    if selected and not authenticated:
        raise RuntimeError(f"浏览器会话 {selected} 未登录 1688，请在对应 Chrome 窗口登录后检查。")
    if len(authenticated) == 1:
        return authenticated[0]
    if not authenticated:
        raise RuntimeError("没有已登录 1688 的浏览器会话。请在 Chrome 中登录 1688，再点“检查并启动”刷新会话。")
    raise RuntimeError("检测到多个已登录的 1688 会话。请在“浏览器会话”列表中选定一个后重试。")


def check_1688_profiles(executable: str | None = None) -> dict[str, Any]:
    profiles = list_1688_profiles(executable)
    authenticated = []
    for profile in profiles:
        try:
            _run(["1688", "whoami", "--format", "json"], timeout=25, executable=executable, profile=profile)
            authenticated.append(profile)
        except RuntimeError:
            pass
    return {"profiles": profiles, "authenticated": authenticated}


def _json_payload(text: str) -> Any:
    # OpenCLI can print runtime notices before its JSON result.
    decoder = json.JSONDecoder()
    for i, char in enumerate(text):
        if char not in "[{": continue
        try:
            value, _ = decoder.raw_decode(text[i:])
            return value
        except ValueError:
            continue
    return {}


def download_1688(url: str, output_root: Path, executable: str | None = None,
                  profile: str | None = None) -> dict[str, Any]:
    if not re.search(r"(^|\.)1688\.com|detail\.1688\.com", url, re.I):
        raise ValueError("请输入有效的 1688 商品或店铺链接。")
    output_root = output_root.expanduser().resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    details_text = _run(["1688", "item", url, "--format", "json"], timeout=240, executable=executable, profile=profile)
    details = _json_payload(details_text)
    if isinstance(details, list): details = details[0] if details else {}
    if not isinstance(details, dict): details = {}
    details["_source_url"] = url
    title = str(details.get("title") or details.get("name") or "1688 商品").strip()
    supplier = str(details.get("seller_name") or details.get("supplier_name") or "1688供应商").strip()
    match = re.search(r"(?:offer/|offerId=)(\d+)", url, re.I)
    item_id = str(details.get("offer_id") or (match.group(1) if match else ""))
    product_name = safe_stem(title)
    supplier_dir = output_root / safe_stem(supplier)
    product_dir = supplier_dir / product_name
    if product_dir.exists():
        # Never replace an earlier manually completed folder on a repeated download.
        suffix = item_id or __import__("hashlib").sha256(url.encode()).hexdigest()[:8]
        product_dir = supplier_dir / f"{product_name}_{suffix}"
    product_dir.mkdir(parents=True, exist_ok=True)
    assets_dir = product_dir / "原始资料" / "1688下载素材"
    assets_dir.mkdir(parents=True, exist_ok=True)
    _run(["1688", "download", url, "--output", str(assets_dir), "--format", "json"], timeout=900, executable=executable, profile=profile)
    (product_dir / "1688商品信息.json").write_text(json.dumps(details, ensure_ascii=False, indent=2), encoding="utf-8")
    (product_dir / "原始资料" / "1688商品信息.json").write_text(json.dumps(details, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"supplier": supplier, "title": title, "product_dir": str(product_dir), "source_dir": str(product_dir), "details": details, "source_product_id": item_id}


def _first_dict(value: Any) -> dict[str, Any]:
    if isinstance(value, list):
        return next((item for item in value if isinstance(item, dict)), {})
    return value if isinstance(value, dict) else {}


def _write_1688_archive_details(destination: Path, details: dict[str, Any], source_url: str,
                                item_error: str = "", asset_error: str = "") -> list[dict[str, str]]:
    """Write source attributes and categorized packaging/certificate records to a standard archive."""
    attributes = details.get("visible_attributes")
    if not isinstance(attributes, list):
        attributes = []
    attributes = [row for row in attributes if isinstance(row, dict)]
    packaging_pattern = re.compile(r"包装|装箱|箱规|外箱|净含量|毛重|净重|尺寸|规格|容量|重量|包装方式", re.I)
    certificate_pattern = re.compile(r"证书|许可证|资质|认证|检测|检验|备案|生产许可|卫生许可", re.I)
    packaging = [row for row in attributes if packaging_pattern.search(str(row.get("key") or ""))]
    certificates = [row for row in attributes if certificate_pattern.search(str(row.get("key") or ""))]
    packaging_blocks = details.get("packaging_info") or details.get("packaging") or {
        key: details.get(key) for key in ("customization_text", "moq_text", "delivery_days_text", "origin_place")
        if details.get(key) not in (None, "")
    }
    certificate_blocks = details.get("qualification_certificates") or details.get("certificates") or []
    record_base = {"source": "1688 item / assets", "source_url": source_url,
                   "collected_at": details.get("fetched_at", ""), "offer_id": details.get("offer_id", "")}
    specs_dir = destination / "规格与包装"
    qualification_dir = destination / "资质与说明"
    specs_dir.mkdir(parents=True, exist_ok=True)
    qualification_dir.mkdir(parents=True, exist_ok=True)
    generated: list[tuple[Path, dict[str, Any]]] = [
        (specs_dir / "商品属性.json", {**record_base, "attributes": attributes,
                                       "price_tiers": details.get("price_tiers", []),
                                       "seller": {key: details.get(key) for key in ("seller_name", "seller_url", "shop_name", "shop_id")}}),
        (specs_dir / "包装信息.json", {**record_base, "attributes": packaging, "packaging_info": packaging_blocks}),
        (qualification_dir / "商品资质证书.json", {**record_base, "attributes": certificates,
                                                    "certificates": certificate_blocks,
                                                    "collection_note": "只记录商品页 item 可见属性和下载素材；若证书为单独页面附件且未出现在属性中，请在原商品页核验。"}),
    ]
    if item_error:
        generated[0][1]["collection_warning"] = f"商品属性读取失败：{item_error}"
    if asset_error:
        generated[2][1]["collection_warning"] = f"图片/视频下载失败：{asset_error}"
    metadata: list[dict[str, str]] = []
    manifest_path = destination / "产品档案.json"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        manifest = {}
    source_files = manifest.setdefault("source_files", [])
    for path, payload in generated:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        entry = {"path": path.relative_to(destination).as_posix(), "source": "1688 item", "sha256": file_sha256(path)}
        source_files[:] = [row for row in source_files if row.get("path") != entry["path"]]
        source_files.append(entry)
        metadata.append(entry)
    if manifest:
        manifest["source_files"] = source_files
        manifest["updated_at"] = __import__("datetime").datetime.now().astimezone().isoformat(timespec="seconds")
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return metadata


def _sort_1688_archive_images(destination: Path) -> None:
    """Separate SKU and detail-page assets into their standard archive sections."""
    manifest_path = destination / "产品档案.json"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return
    changed = False
    for image in manifest.get("images", []):
        source = str(image.get("source") or "")
        if "1688下载素材/" not in source:
            continue
        lowered = Path(source).name.lower()
        if "_sku_" in lowered:
            section = Path("规格与包装") / "SKU图片"
        elif "_detail_" in lowered or "_other_" in lowered:
            section = Path("资质与说明") / "商品详情图片"
        else:
            continue
        current = destination / str(image.get("path") or "")
        if not current.is_file():
            continue
        target_dir = destination / section
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / current.name
        if target.exists():
            target = target.with_name(f"{target.stem}_{file_sha256(current)[:8]}{target.suffix}")
        shutil.move(str(current), str(target))
        image["path"] = target.relative_to(destination).as_posix()
        changed = True
    if changed:
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def _merge_1688_archive_media(destination: Path, info: dict[str, Any]) -> None:
    """Add newly discovered assets to an existing reviewed archive without replacing it."""
    manifest_path = destination / "产品档案.json"
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return
    known_images = {str(row.get("sha256") or "") for row in manifest.get("images", [])}
    known_videos = {str(row.get("sha256") or "") for row in manifest.get("videos", [])}
    for source in info.get("images", []):
        digest = file_sha256(source)
        if digest in known_images:
            continue
        source_relative = source.relative_to(Path(info["folder"])).as_posix()
        filename = f"{safe_stem(source.stem)}{source.suffix.lower()}"
        lowered = source.name.lower()
        if "_sku_" in lowered:
            target_dir = destination / "规格与包装" / "SKU图片"
            role = "sku"
        elif "_detail_" in lowered or "_other_" in lowered:
            target_dir = destination / "资质与说明" / "商品详情图片"
            role = "detail"
        elif not any(row.get("role") == "main" for row in manifest.get("images", [])):
            target_dir = destination / "图片"
            filename = f"主图{source.suffix.lower()}"
            role = "main"
        else:
            target_dir = destination / "图片" / "副图"
            role = "gallery"
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / filename
        if target.exists():
            target = target.with_name(f"{target.stem}_{digest[:8]}{target.suffix}")
        shutil.copy2(source, target)
        manifest.setdefault("images", []).append({"path": target.relative_to(destination).as_posix(),
                                                   "role": role, "source": source_relative, "sha256": digest})
        known_images.add(digest)
    for source in info.get("videos", []):
        digest = file_sha256(source)
        if digest in known_videos:
            continue
        target_dir = destination / "视频"
        target_dir.mkdir(parents=True, exist_ok=True)
        target = target_dir / f"{safe_stem(source.stem)}{source.suffix.lower()}"
        if target.exists():
            target = target.with_name(f"{target.stem}_{digest[:8]}{target.suffix}")
        shutil.copy2(source, target)
        manifest.setdefault("videos", []).append({"path": target.relative_to(destination).as_posix(),
                                                   "source": source.relative_to(Path(info["folder"])).as_posix(), "sha256": digest})
        known_videos.add(digest)
    manifest["updated_at"] = __import__("datetime").datetime.now().astimezone().isoformat(timespec="seconds")
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def download_1688_shop_products(url: str, output_root: Path, executable: str | None = None,
                                profile: str | None = None, max_pages: int = 100,
                                on_progress: Any | None = None) -> dict[str, Any]:
    """Resolve a 1688 factory/store page and run shop-products to build per-item folders."""
    if not re.search(r"(^|\.)1688\.com", url, re.I):
        raise ValueError("请输入有效的 1688 工厂或店铺链接。")
    if not 1 <= int(max_pages) <= 100:
        raise ValueError("shop-products 最大页数必须在 1–100 之间。")
    output_root = output_root.expanduser().resolve()
    output_root.mkdir(parents=True, exist_ok=True)

    if "/factory/" in url.lower():
        if on_progress:
            on_progress("正在解析工厂主页和店铺商品列表")
        store_data = _first_dict(_json_payload(_run(
            ["1688", "store", url, "--format", "json"], timeout=240,
            executable=executable, profile=profile)))
        offerlist_url = str(store_data.get("store_url") or store_data.get("shop_url") or "").strip()
        if not offerlist_url:
            raise RuntimeError("1688 工厂主页未返回店铺商品列表地址，暂时无法运行 shop-products。")
        store_name = str(store_data.get("store_name") or store_data.get("company_name") or "1688 店铺").strip()
        shop_id = str(store_data.get("shop_id") or "")
    else:
        offerlist_url = url
        store_name = "1688 店铺"
        shop_id = ""
    if "/page/offerlist.htm" not in offerlist_url.lower():
        offerlist_url = offerlist_url.rstrip("/") + "/page/offerlist.htm"

    supplier_dir = output_root / safe_stem(store_name)
    supplier_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = supplier_dir / f"1688店铺商品清单_{__import__('datetime').datetime.now():%Y%m%d_%H%M%S}.xlsx"
    if on_progress:
        on_progress("正在用 1688 shop-products 抓取商品列表和主图")
    started_at = time.time()
    output = _run(["1688", "shop-products", offerlist_url, "--maxPages", str(int(max_pages)),
                   "--output", str(workbook_path), "--format", "json"], timeout=1200,
                  executable=executable, profile=profile, progress_file=supplier_dir,
                  on_progress=on_progress)
    if not workbook_path.is_file():
        # OpenCLI may name the export after the shop id; accept the newest workbook
        # in this isolated supplier folder, but never scan unrelated output folders.
        candidates = sorted((path for path in supplier_dir.glob("*.xlsx") if path.stat().st_mtime >= started_at - 1),
                            key=lambda p: p.stat().st_mtime, reverse=True)
        if candidates:
            workbook_path = candidates[0]
        else:
            raise RuntimeError(f"shop-products 已结束，但没有找到导出的 Excel。\n{output[-2000:]}")

    from openpyxl import load_workbook

    workbook = load_workbook(workbook_path, read_only=False, data_only=True)
    products: list[dict[str, Any]] = []
    try:
        sheet = workbook.active
        total_rows = max(1, sheet.max_row - 1)
        embedded_images: dict[int, list[Any]] = {}
        for image in sheet._images:
            row_number = image.anchor._from.row + 1
            embedded_images.setdefault(row_number, []).append(image)
        staging_root = Path(tempfile.mkdtemp(prefix="xudaai-1688-shop-"))
        try:
          for row_number in range(2, sheet.max_row + 1):
            offer_id = str(sheet.cell(row_number, 1).value or "").strip()
            detail_url = str(sheet.cell(row_number, 2).value or "").strip()
            title = str(sheet.cell(row_number, 3).value or "").strip()
            if not title and not offer_id and not detail_url:
                continue
            title = title or f"1688 商品 {offer_id or row_number}"
            offer_id = offer_id or (re.search(r"offer/(\d+)", detail_url).group(1) if re.search(r"offer/(\d+)", detail_url) else str(row_number))
            price = sheet.cell(row_number, 4).value
            sales = sheet.cell(row_number, 5).value
            moq = sheet.cell(row_number, 6).value
            image_url = str(sheet.cell(row_number, 8).value or "").strip()
            page = sheet.cell(row_number, 9).value
            source_url = detail_url or f"https://detail.1688.com/offer/{offer_id}.html"
            details = {"offer_id": offer_id, "title": title, "seller_name": store_name,
                       "supplier_name": store_name, "price": price, "moq": moq,
                       "sales_text": sales, "image_url": image_url, "page": page,
                       "_source_url": source_url, "store_url": offerlist_url,
                       "shop_id": shop_id}
            stage_dir = staging_root / f"{safe_stem(title)}_{safe_stem(offer_id)}"
            assets_dir = stage_dir / "1688下载素材"
            assets_dir.mkdir(parents=True, exist_ok=True)
            item_error = ""
            asset_error = ""
            item_payload: dict[str, Any] = {}
            item_number = len(products) + 1
            if on_progress:
                on_progress(f"商品处理 {item_number}/{total_rows}：读取 {title} 的属性和包装信息")
            try:
                item_payload = _first_dict(_json_payload(_run(
                    ["1688", "item", offer_id, "--format", "json"], timeout=240,
                    executable=executable, profile=profile)))
            except (RuntimeError, subprocess.TimeoutExpired) as exc:
                item_error = str(exc)
            details.update(item_payload)
            details.update({"offer_id": offer_id, "title": item_payload.get("title") or title,
                            "seller_name": item_payload.get("seller_name") or store_name,
                            "supplier_name": store_name, "price": price, "moq": moq,
                            "sales_text": sales, "image_url": image_url, "page": page,
                            "_source_url": source_url, "store_url": offerlist_url, "shop_id": shop_id})
            if on_progress:
                on_progress(f"商品处理 {item_number}/{total_rows}：使用 1688 assets 识别并下载图片/视频")
            try:
                # The download command uses the same assets extractor internally, then persists
                # its main/SKU/detail/other images and videos beneath assets_dir/<offer_id>/.
                _run(["1688", "download", offer_id, "--output", str(assets_dir), "--format", "json"],
                     timeout=900, executable=executable, profile=profile)
            except (RuntimeError, subprocess.TimeoutExpired) as exc:
                asset_error = str(exc)
            metadata = json.dumps(details, ensure_ascii=False, indent=2)
            (stage_dir / "1688商品信息.json").write_text(metadata, encoding="utf-8")
            # Keep the package within the core's 250-file limit, preserving video files first.
            media_files = sorted((path for path in assets_dir.rglob("*") if path.is_file() and
                                  path.suffix.lower() in {".jpg", ".jpeg", ".png", ".webp", ".gif", ".mp4", ".m3u8", ".mov", ".webm"}),
                                 key=lambda path: (path.suffix.lower() not in {".mp4", ".m3u8", ".mov", ".webm"}, path.name.lower()))
            for excess in media_files[max(0, 240 - len(embedded_images.get(row_number, []))):]:
                try:
                    excess.unlink()
                except OSError:
                    pass
            for image_index, image in enumerate(embedded_images.get(row_number, []), 1):
                extension = str(getattr(image, "format", "jpeg") or "jpeg").lower().replace("jpg", "jpeg")
                extension = extension if extension in {"jpeg", "png", "webp", "gif"} else "jpeg"
                image_path = assets_dir / f"主图_{image_index}.{extension}"
                image_path.write_bytes(image._data())
            info = scan_product_folder(stage_dir, max_files=30000)
            stable_source = detail_url or f"1688:{shop_id}:{offer_id}"
            product_id = str(uuid.uuid5(uuid.NAMESPACE_URL, stable_source))
            product_code = "P-" + uuid.UUID(product_id).hex[:12].upper()
            values = {"name": title, "supplier_name": store_name,
                      "cost_price": str(price) if price is not None else "",
                      "unit_price": str(price) if price is not None else "",
                      "currency": "CNY", "moq": str(moq) if moq is not None else "",
                      "notes": f"店铺列表销量：{sales}" if sales not in (None, "") else "",
                      "source_type": "1688/OpenCLI", "source_url": source_url,
                      "source_product_id": offer_id, "product_id": product_id,
                      "product_code": product_code,
                      "field_sources": {"name": "1688 shop-products 商品标题",
                                        "supplier_name": "1688 店铺信息",
                                        "cost_price": "1688 shop-products 商品列表价格",
                                        "moq": "1688 shop-products 商品列表起订量"}}
            downloaded_main = next((path for path in info["images"] if re.search(r"_main_0?1\.", path.name, re.I)), None)
            main_path = downloaded_main or (info["images"][0] if info["images"] else None)
            main_image = main_path.relative_to(stage_dir).as_posix() if main_path else ""
            destination = supplier_dir / f"{product_code}-{safe_stem(title)}"
            manifest_path = destination / "产品档案.json"
            if manifest_path.is_file():
                try:
                    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                except (OSError, ValueError):
                    manifest = {}
                if manifest.get("product_id") != product_id:
                    raise FileExistsError(f"标准档案目录已存在且来源不匹配，未覆盖：{destination}")
                # Re-importing the same item must preserve fields already reviewed by the user.
                _merge_1688_archive_media(destination, info)
            else:
                destination, manifest = make_product_package(info, None, output_root, values, main_image)
                _sort_1688_archive_images(destination)
            _write_1688_archive_details(destination, details, source_url, item_error, asset_error)
            products.append({"supplier": store_name, "title": title, "product_dir": str(destination),
                             "source_dir": str(destination), "details": details,
                             "manifest": manifest, "source_product_id": offer_id,
                             "asset_count": len([path for path in info["images"] + info["videos"]]),
                             "item_error": item_error, "asset_error": asset_error})
            if on_progress:
                on_progress(f"商品处理 {item_number}/{total_rows} 完成：图片 {len(info['images'])}，视频 {len(info['videos'])}")
        finally:
            shutil.rmtree(staging_root, ignore_errors=True)
    finally:
        workbook.close()
    if not products:
        raise RuntimeError(f"shop-products 没有找到商品行，请检查店铺登录状态和页面。\n{output[-2000:]}")
    return {"supplier": store_name, "title": f"{store_name} 店铺商品", "product_dir": str(supplier_dir),
            "source_dir": str(supplier_dir), "details": {"store_url": offerlist_url},
            "products": products, "product_count": len(products), "standard_archives": True, "workbook": str(workbook_path),
            "source_product_id": shop_id}
