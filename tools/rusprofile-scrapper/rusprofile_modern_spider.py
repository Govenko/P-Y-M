#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Modern Rusprofile spider for current list pages."""

from __future__ import annotations

import re
from urllib.parse import urljoin

import scrapy


class ModernRusprofileSpider(scrapy.Spider):
    name = "rusprofile_modern"

    CODE_PAGE_RE = re.compile(r"/codes/(\d+)(?:/(\d+))?")

    def __init__(self, ids=None, max_pages=0, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.ids = [str(value).strip() for value in (ids or []) if str(value).strip()]
        self.max_pages = int(max_pages or 0)  # 0 = без ограничения
        self.seen_keys = set()
        self.start_urls = [f"https://www.rusprofile.ru/codes/{value}" for value in self.ids]

    def parse(self, response):
        match = self.CODE_PAGE_RE.search(response.url)
        code = match.group(1) if match else ""
        page_no = int(match.group(2) or 1) if match else 1
        new_items = 0

        for item in response.css(".list-element, div.company-item"):
            parsed = self.parse_list_item(item, response)
            key = parsed.get("ogrn") or parsed.get("inn") or (parsed.get("name", "") + parsed.get("source_url", ""))
            if not key or key in self.seen_keys:
                continue
            self.seen_keys.add(key)
            new_items += 1
            if parsed.get("ogrn"):
                yield parsed
                continue

            href = item.css("a.list-element__title::attr(href), div.company-item__title a::attr(href)").get()
            if href:
                yield response.follow(href, callback=self.parse_company_page)

        # Каталог rusprofile листается прямыми URL /codes/{code}/{n}.
        # Останавливаемся, когда страница не дала ни одной новой компании.
        if new_items > 0 and (self.max_pages <= 0 or page_no < self.max_pages):
            next_no = page_no + 1
            yield scrapy.Request(
                f"https://www.rusprofile.ru/codes/{code}/{next_no}",
                callback=self.parse,
                meta={"code": code, "page_no": next_no},
            )
        else:
            self.logger.info("Code %s finished at page %s (new items: %s)", code, page_no, new_items)

    def parse_list_item(self, item, response) -> dict:
        text = normalize(" ".join(item.css("::text").getall()))
        title = first_clean(item.css("a.list-element__title::text, div.company-item__title a::text").getall())
        href = item.css("a.list-element__title::attr(href), div.company-item__title a::attr(href)").get() or ""
        return {
            "name": title,
            "inn": pick(text, r"ИНН:\s*(\d{10}|\d{12})"),
            "ogrn": pick(text, r"ОГРН:\s*(\d{13}|\d{15})"),
            "okpo": "",
            "current_status": pick_status(text),
            "reg_date": normalize_date(pick(text, r"Дата регистрации:\s*(\d{2}\.\d{2}\.\d{4})")),
            "auth_capital": "",
            "address": extract_address_from_list_text(text),
            "source_url": urljoin(response.url, href) if href else response.url,
            "source_code_url": response.url,
        }

    def parse_company_page(self, response):
        anketa_text = normalize(" ".join(response.css("#anketa ::text").getall()))
        all_text = normalize(" ".join(response.css("body ::text").getall()))
        data = {
            "name": first_clean(response.css("div.company-name::text, .company-header__title::text, h1::text").getall()),
            "inn": first_clean(response.css("#clip_inn::text").getall()) or pick(all_text, r"ИНН/КПП\s*(\d{10}|\d{12})"),
            "ogrn": first_clean(response.css("#clip_ogrn::text").getall()),
            "okpo": first_clean(response.css("#clip_okpo::text").getall()),
            "current_status": pick_status(first_clean(response.css("div.company-status::text, .company-header__status::text").getall()) or all_text),
            "reg_date": normalize_date(pick(anketa_text, r"Дата регистрации\s*(\d{2}\.\d{2}\.\d{4})")),
            "auth_capital": normalize_money(pick(anketa_text, r"Уставный капитал\s*([\d\s]+(?:руб\.|₽)?)")),
            "address": pick(anketa_text, r"Юридический адрес\s*(.+?)(?:Руководитель|Основной вид|Все реквизиты|$)"),
            "source_url": response.url,
            "source_code_url": "",
        }
        if data["ogrn"]:
            yield data


def normalize(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def first_clean(values) -> str:
    for value in values or []:
        cleaned = normalize(value)
        if cleaned:
            return cleaned
    return ""


def pick(text: str, pattern: str) -> str:
    match = re.search(pattern, text or "", re.IGNORECASE)
    return normalize(match.group(1)) if match else ""


def normalize_date(value: str) -> str:
    if not value:
        return ""
    parts = value.split(".")
    if len(parts) == 3:
        return "".join(reversed(parts))
    return value


def normalize_money(value: str) -> str:
    digits = "".join(re.findall(r"\d+", value or ""))
    return digits


def pick_status(text: str) -> str:
    text = normalize(text)
    known = (
        "Действующая организация",
        "Ликвидирована",
        "В процессе ликвидации",
        "Реорганизована",
    )
    for status in known:
        if status.lower() in text.lower():
            return status
    return ""


def extract_address_from_list_text(text: str) -> str:
    inn_match = re.search(r"\sИНН:\s*(?:\d{10}|\d{12})", text or "")
    if not inn_match:
        return ""
    before_inn = text[: inn_match.start()].strip()
    match = re.search(r"(\d{6}\s*,?\s+.+)$", before_inn)
    return normalize(match.group(1)) if match else ""
