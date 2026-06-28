#!/usr/bin/env python3
"""
PaddleOCR API batch processor
Usage: python3 paddle_ocr.py <input_dir> <output_dir> [--token TOKEN]
"""

import argparse
import json
import os
import sys
import time
import requests

JOB_URL = "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs"
MODEL = "PaddleOCR-VL-1.6"
POLL_INTERVAL = 5  # seconds


def submit_job(file_path: str, token: str) -> str:
    headers = {"Authorization": f"bearer {token}"}
    optional_payload = {
        "useDocOrientationClassify": False,
        "useDocUnwarping": False,
        "useChartRecognition": False,
    }
    data = {
        "model": MODEL,
        "optionalPayload": json.dumps(optional_payload),
    }
    with open(file_path, "rb") as f:
        resp = requests.post(JOB_URL, headers=headers, data=data, files={"file": f}, timeout=60)

    if resp.status_code != 200:
        raise RuntimeError(f"提交失败 ({resp.status_code}): {resp.text}")
    return resp.json()["data"]["jobId"]


def poll_job(job_id: str, token: str, label: str) -> str:
    headers = {"Authorization": f"bearer {token}"}
    while True:
        resp = requests.get(f"{JOB_URL}/{job_id}", headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()["data"]
        state = data["state"]

        if state == "pending":
            print(f"    等待中...", flush=True)
        elif state == "running":
            try:
                prog = data["extractProgress"]
                print(f"    识别中 {prog['extractedPages']}/{prog['totalPages']} 页", flush=True)
            except KeyError:
                print(f"    识别中...", flush=True)
        elif state == "done":
            prog = data.get("extractProgress", {})
            print(f"    完成，共 {prog.get('extractedPages', '?')} 页", flush=True)
            return data["resultUrl"]["jsonUrl"]
        elif state == "failed":
            raise RuntimeError(f"任务失败: {data.get('errorMsg', '未知错误')}")

        time.sleep(POLL_INTERVAL)


def download_markdown(jsonl_url: str) -> str:
    resp = requests.get(jsonl_url, timeout=60)
    resp.raise_for_status()
    lines = resp.text.strip().split("\n")

    pages = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        result = json.loads(line)["result"]
        for layout in result.get("layoutParsingResults", []):
            pages.append(layout["markdown"]["text"])

    if len(pages) > 1:
        return "\n\n---\n\n".join(pages)
    elif pages:
        return pages[0]
    return ""


def process_directory(input_dir: str, output_dir: str, token: str):
    pdfs = sorted([f for f in os.listdir(input_dir) if f.lower().endswith(".pdf")])
    if not pdfs:
        print(f"未找到 PDF 文件: {input_dir}")
        return

    os.makedirs(output_dir, exist_ok=True)
    total = len(pdfs)
    success, failed = 0, []

    for i, pdf_name in enumerate(pdfs, 1):
        pdf_path = os.path.join(input_dir, pdf_name)
        stem = os.path.splitext(pdf_name)[0]
        md_path = os.path.join(output_dir, f"{stem}.md")

        print(f"[{i}/{total}] {stem}", flush=True)

        try:
            job_id = submit_job(pdf_path, token)
            print(f"    任务ID: {job_id}", flush=True)
            jsonl_url = poll_job(job_id, token, stem)
            markdown = download_markdown(jsonl_url)

            if not markdown.startswith(f"# {stem}"):
                markdown = f"# {stem}\n\n{markdown}"

            with open(md_path, "w", encoding="utf-8") as f:
                f.write(markdown)

            print(f"    ✅ 已保存 → {stem}.md", flush=True)
            success += 1

        except Exception as e:
            print(f"    ❌ 失败: {e}", flush=True)
            failed.append((stem, str(e)))

    print(f"\n{'='*50}")
    print(f"完成: {success}/{total} 个文件")
    if failed:
        print(f"失败 ({len(failed)} 个):")
        for name, err in failed:
            print(f"  - {name}: {err}")
    print(f"输出目录: {output_dir}")


def main():
    parser = argparse.ArgumentParser(description="PaddleOCR API 批量处理 PDF")
    parser.add_argument("input_dir", help="PDF 所在目录")
    parser.add_argument("output_dir", help="Markdown 输出目录")
    parser.add_argument("--token", help="API Token（也可用环境变量 PADDLEOCR_TOKEN）")
    args = parser.parse_args()

    token = args.token or os.environ.get("PADDLEOCR_TOKEN")
    if not token:
        print("错误: 请通过 --token 参数或 PADDLEOCR_TOKEN 环境变量提供 API Token", file=sys.stderr)
        sys.exit(1)

    process_directory(args.input_dir, args.output_dir, token)


if __name__ == "__main__":
    main()
