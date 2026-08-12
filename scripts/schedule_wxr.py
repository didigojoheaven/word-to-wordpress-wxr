#!/usr/bin/env python3
"""Convert a WordPress WXR draft file into chronologically scheduled posts."""

from __future__ import annotations

import argparse
import re
import sys
from datetime import datetime, timedelta, timezone
from email.utils import format_datetime
from pathlib import Path
from xml.etree import ElementTree as etree
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

WP_NS = "http://wordpress.org/export/1.2/"
CONTENT_NS = "http://purl.org/rss/1.0/modules/content/"
DC_NS = "http://purl.org/dc/elements/1.1/"
EXCERPT_NS = "http://wordpress.org/export/1.2/excerpt/"

etree.register_namespace("wp", WP_NS)
etree.register_namespace("content", CONTENT_NS)
etree.register_namespace("dc", DC_NS)
etree.register_namespace("excerpt", EXCERPT_NS)


def wp_tag(name: str) -> str:
    return f"{{{WP_NS}}}{name}"


def parse_start(value: str, timezone_name: str) -> datetime:
    try:
        zone = ZoneInfo(timezone_name)
    except ZoneInfoNotFoundError as error:
        raise ValueError(f"Unknown timezone: {timezone_name}") from error
    try:
        start = datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError as error:
        raise ValueError("Use YYYY-MM-DD HH:MM for --start") from error
    if start.tzinfo is None:
        start = start.replace(tzinfo=zone)
    else:
        start = start.astimezone(zone)
    if start <= datetime.now(zone):
        raise ValueError("--start must be in the future to avoid immediate publication")
    return start


def article_key(item: etree.Element) -> tuple[object, ...]:
    post_name = item.findtext(wp_tag("post_name")) or ""
    return tuple(int(part) if part.isdigit() else part.casefold() for part in re.split(r"(\d+)", post_name))


def set_text(item: etree.Element, tag: str, value: str) -> None:
    node = item.find(tag)
    if node is None:
        node = etree.SubElement(item, tag)
    node.text = value


def schedule(input_path: Path, output_path: Path, start: datetime, interval_minutes: int) -> int:
    if input_path.resolve() == output_path.resolve():
        raise ValueError("--output must be a new file; do not overwrite the draft import")
    tree = etree.parse(input_path)
    channel = tree.find("./channel")
    if channel is None:
        raise ValueError("Not a WordPress WXR file: channel not found")
    items = channel.findall("item")
    if not items:
        raise ValueError("No post items found in WXR")

    items.sort(key=article_key)
    for item in channel.findall("item"):
        channel.remove(item)
    for item in items:
        channel.append(item)

    for index, item in enumerate(items):
        publish_at = start + timedelta(minutes=index * interval_minutes)
        set_text(item, "pubDate", format_datetime(publish_at.astimezone(timezone.utc), usegmt=True))
        set_text(item, wp_tag("post_date"), publish_at.strftime("%Y-%m-%d %H:%M:%S"))
        set_text(item, wp_tag("post_date_gmt"), publish_at.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"))
        set_text(item, wp_tag("status"), "future")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    tree.write(output_path, encoding="utf-8", xml_declaration=True)
    validate(output_path, start, interval_minutes)
    return len(items)


def validate(path: Path, start: datetime, interval_minutes: int) -> None:
    root = etree.parse(path).getroot()
    items = root.findall("./channel/item")
    if not items:
        raise ValueError("Validation failed: no post items")
    for index, item in enumerate(items):
        if item.findtext(wp_tag("status")) != "future":
            raise ValueError(f"Validation failed: item {index + 1} is not scheduled")
        expected = start + timedelta(minutes=index * interval_minutes)
        local_value = item.findtext(wp_tag("post_date")) or ""
        gmt_value = item.findtext(wp_tag("post_date_gmt")) or ""
        if local_value != expected.strftime("%Y-%m-%d %H:%M:%S"):
            raise ValueError(f"Validation failed: incorrect local time for item {index + 1}")
        if gmt_value != expected.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M:%S"):
            raise ValueError(f"Validation failed: incorrect UTC time for item {index + 1}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Existing draft WXR/XML file")
    parser.add_argument("--output", required=True, help="New scheduled WXR/XML file")
    parser.add_argument("--start", required=True, help="First publication time, e.g. 2026-07-31 15:35")
    parser.add_argument("--interval-minutes", required=True, type=int)
    parser.add_argument("--timezone", default="Asia/Shanghai")
    args = parser.parse_args()
    if args.interval_minutes <= 0:
        parser.error("--interval-minutes must be greater than 0")
    args.start = parse_start(args.start, args.timezone)
    return args


def main() -> int:
    args = parse_args()
    try:
        count = schedule(Path(args.input), Path(args.output), args.start, args.interval_minutes)
        print(f"Created and validated {count} scheduled posts: {args.output}")
        return 0
    except Exception as error:
        print(f"Failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
