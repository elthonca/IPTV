#!/usr/bin/env python3
"""
Read pluto_channels.xlsx, rebuild the original M3U files (one per country sheet)
and upload them to a GitHub repository via the REST API.

Requirements:  pip install openpyxl
Token:         set the GITHUB_TOKEN environment variable (a Personal Access Token
               with "repo" scope, or a fine-grained token with
               "Contents: Read and write" on the target repository).

Examples:
  # Test locally only (writes ./m3u_out/pluto_br.m3u etc., uploads nothing)
  python upload_m3u_to_github.py --dry-run

  # Upload to github.com/youruser/iptv  (folder "playlists" on branch main)
  export GITHUB_TOKEN=ghp_xxx          # PowerShell: $env:GITHUB_TOKEN="ghp_xxx"
  python upload_m3u_to_github.py --repo youruser/iptv --folder playlists
"""
import argparse
import base64
import json
import os
import sys
import urllib.error
import urllib.request

from openpyxl import load_workbook

EPG_URL = "https://github.com/matthuisman/i.mjh.nz/raw/master/PlutoTV/{cc}.xml.gz"
COUNTRY_SHEETS = ["BR", "CA", "US"]  # "ALL" is skipped to avoid duplicates


def build_m3u(ws, country):
    """Rebuild one M3U playlist from a worksheet.

    Columns: Channel Name | Group | TVG ID | Logo URL | Stream URL | Country
    """
    lines = [f'#EXTM3U url-tvg="{EPG_URL.format(cc=country.lower())}"']
    for name, group, tvg_id, logo, url, *_ in ws.iter_rows(min_row=2, values_only=True):
        if not url:  # skip blank rows
            continue
        lines.append(
            f'#EXTINF:-1 tvg-id="{tvg_id or ""}" tvg-logo="{logo or ""}" '
            f'group-title="{group or ""}",{name or ""}'
        )
        lines.append(str(url).strip())
    return "\n".join(lines) + "\n"


def github_request(method, url, token, payload=None):
    req = urllib.request.Request(
        url,
        method=method,
        data=json.dumps(payload).encode("utf-8") if payload is not None else None,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
            "User-Agent": "m3u-uploader",
        },
    )
    with urllib.request.urlopen(req) as resp:
        return json.load(resp)


def upload_file(repo, branch, path, content, token, message):
    api = f"https://api.github.com/repos/{repo}/contents/{path}"
    # If the file already exists we must send its current SHA to update it.
    sha = None
    try:
        sha = github_request("GET", f"{api}?ref={branch}", token)["sha"]
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
    payload = {
        "message": message,
        "content": base64.b64encode(content.encode("utf-8")).decode("ascii"),
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha
    result = github_request("PUT", api, token, payload)
    return ("updated" if sha else "created"), result["content"]["html_url"]


def main():
    p = argparse.ArgumentParser(description="Spreadsheet -> M3U -> GitHub")
    p.add_argument("--xlsx", default="./Source/pluto_channels.xlsx", help="input spreadsheet")
    p.add_argument("--repo", help="GitHub repo as owner/name")
    p.add_argument("--branch", default="main")
    p.add_argument("--folder", default="", help="folder inside the repo (optional)")
    p.add_argument("--message", default="Update Pluto TV playlists")
    p.add_argument("--out", default="m3u_out", help="local output folder")
    p.add_argument("--dry-run", action="store_true", help="only write files locally")
    args = p.parse_args()

    token = os.environ.get("GITHUB_TOKEN")
    if not args.dry_run:
        if not args.repo:
            sys.exit("Error: --repo owner/name is required (or use --dry-run).")
        if not token:
            sys.exit("Error: set the GITHUB_TOKEN environment variable.")

    wb = load_workbook(args.xlsx, read_only=True, data_only=True)
    os.makedirs(args.out, exist_ok=True)

    for cc in COUNTRY_SHEETS:
        if cc not in wb.sheetnames:
            print(f"Sheet {cc} not found, skipping.")
            continue
        content = build_m3u(wb[cc], cc)
        filename = f"pluto_{cc.lower()}.m3u"
        local_path = os.path.join(args.out, filename)
        with open(local_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(content)
        n = content.count("#EXTINF")
        print(f"{filename}: {n} channels written to {local_path}")

        if args.dry_run:
            continue
        repo_path = f"{args.folder.strip('/')}/{filename}".lstrip("/")
        try:
            action, link = upload_file(
                args.repo, args.branch, repo_path, content, token, args.message
            )
            print(f"  -> {action}: {link}")
        except urllib.error.HTTPError as e:
            print(f"  -> FAILED ({e.code}): {e.read().decode('utf-8', 'ignore')[:300]}")


if __name__ == "__main__":
    main()