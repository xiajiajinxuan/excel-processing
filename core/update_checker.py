# -*- coding: utf-8 -*-
"""
检查更新模块：从 GitHub/Gitee Releases API 获取最新版本信息，比较版本号并返回结果。
使用标准库 urllib，不新增第三方依赖。
"""

import json
import re
import urllib.error
import urllib.parse
import urllib.request
import ssl


def parse_version(version_str):
    """
    将版本字符串规范化为可比较的整数元组。
    支持 v1.0.0、1.0.0、1.0 等格式。
    """
    if not version_str or not isinstance(version_str, str):
        return (0, 0, 0)
    s = version_str.strip().lower()
    if s.startswith("v"):
        s = s[1:]
    parts = re.findall(r"\d+", s)
    if not parts:
        return (0, 0, 0)
    return tuple(int(x) for x in parts[:3])  # 最多取三段


def compare_versions(current_str, latest_str):
    """
    比较两个版本号。仅当 latest 严格大于 current 时返回 True。
    """
    cur = parse_version(current_str)
    lat = parse_version(latest_str)
    return lat > cur


def _fetch_json(url, timeout=10):
    """使用 urllib 请求 URL 并解析 JSON。失败返回 None 并打印/记录原因。"""
    try:
        ctx = ssl.create_default_context()
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            if resp.status != 200:
                return None
            return json.loads(resp.read().decode("utf-8"))
    except (urllib.error.URLError, urllib.error.HTTPError, OSError, json.JSONDecodeError) as e:
        return None


def _get_github_download_url(assets):
    """从 GitHub release 的 assets 中取第一个 .exe 的 browser_download_url。"""
    if not assets:
        return None
    for a in assets:
        name = (a.get("name") or "").lower()
        if name.endswith(".exe"):
            return a.get("browser_download_url")
    return assets[0].get("browser_download_url")


def _get_gitee_download_url(owner, repo, tag, assets_or_attach):
    """
    Gitee 的 release 可能返回 assets 或 attach_files。
    下载接口：GET /api/v5/repos/:owner/:repo/releases/:tag/attach_files/:file_name/download
    若 assets 项中有 browser_download_url 则直接使用，否则用上述格式拼接。
    """
    if not assets_or_attach:
        return None
    tag_clean = tag.lstrip("v")
    for a in assets_or_attach:
        name = (a.get("name") or a.get("filename") or "").strip()
        if not name:
            continue
        if name.lower().endswith(".exe"):
            url = a.get("browser_download_url") or a.get("url")
            if url:
                return url
            # 拼接 Gitee API 下载地址
            return (
                f"https://gitee.com/api/v5/repos/{owner}/{repo}/releases/"
                f"{tag_clean}/attach_files/{urllib.parse.quote(name)}/download"
            )
    a = assets_or_attach[0]
    name = (a.get("name") or a.get("filename") or "").strip()
    url = a.get("browser_download_url") or a.get("url")
    if url:
        return url
    if name:
        return (
            f"https://gitee.com/api/v5/repos/{owner}/{repo}/releases/"
            f"{tag_clean}/attach_files/{urllib.parse.quote(name)}/download"
        )
    return None


def check_update(config, current_version):
    """
    根据 config 中的 update 配置检查是否有新版本。

    config: 完整配置字典（含 update 键）。
    current_version: 当前应用版本号字符串（如 "1.0.0"）。

    返回字典：
        has_new: bool - 是否有新版本
        current: str - 当前版本
        latest: str - 最新版本（若有）
        release_notes: str - 更新说明（若有）
        download_url: str - 下载链接（若有）
        error: str - 若检查失败时的错误描述（未启用、网络错误等）
    """
    result = {
        "has_new": False,
        "current": current_version or "0.0.0",
        "latest": "",
        "release_notes": "",
        "download_url": "",
        "error": "",
    }
    update_cfg = (config or {}).get("update") or {}
    if not update_cfg.get("enabled", False):
        result["error"] = "未启用更新检查"
        return result

    owner = (update_cfg.get("owner") or "").strip()
    repo = (update_cfg.get("repo") or "").strip()
    source = (update_cfg.get("source") or "github").strip().lower()
    tag_prefix = (update_cfg.get("tag_prefix") or "").strip()

    if not owner or not repo:
        result["error"] = "更新配置缺少 owner 或 repo"
        return result

    if source == "gitee":
        url = f"https://gitee.com/api/v5/repos/{owner}/{repo}/releases/latest"
    else:
        url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"

    data = _fetch_json(url)
    if not data:
        result["error"] = "无法获取发布信息（网络或接口异常）"
        return result

    tag_name = (data.get("tag_name") or "").strip()
    if tag_prefix and not tag_name.startswith(tag_prefix):
        result["error"] = "暂无适用于本应用的发布"
        return result
    latest_version = tag_name.lstrip("v").strip()
    if not latest_version:
        result["error"] = "发布信息无效"
        return result

    result["latest"] = latest_version
    result["release_notes"] = (data.get("body") or "").strip() or "无更新说明"

    if source == "github":
        result["download_url"] = _get_github_download_url(data.get("assets") or [])
    else:
        assets = data.get("assets") or data.get("attach_files") or []
        result["download_url"] = _get_gitee_download_url(owner, repo, tag_name, assets)

    if compare_versions(current_version, latest_version):
        result["has_new"] = True
    return result


def download_file(url, save_path, timeout=60):
    """
    将 url 下载到 save_path。使用 urllib。
    成功返回 True，失败返回 False（不抛异常）。
    """
    try:
        ctx = ssl.create_default_context()
        req = urllib.request.Request(url, headers={"Accept": "application/octet-stream"})
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            if resp.status != 200:
                return False
            with open(save_path, "wb") as f:
                while True:
                    chunk = resp.read(8192)
                    if not chunk:
                        break
                    f.write(chunk)
        return True
    except (urllib.error.URLError, urllib.error.HTTPError, OSError):
        return False
