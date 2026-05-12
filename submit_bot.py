"""
自动提交脚本 - gaodun online-consultation/add-new
基于抓包数据自动提交在线咨询记录
"""

import requests
import json
import time
import logging
from datetime import datetime
from typing import Optional

# ─────────────────────────────────────────
# 日志配置
# ─────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("submit_bot.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)


# ─────────────────────────────────────────
# 配置区（按需修改）
# ─────────────────────────────────────────
BASE_URL = "https://apigateway.gaodun.com"
ENDPOINT = "/solon/api/v1/online-consultation/add-new"

# 替换为最新的 JWT Token（从抓包的 authentication 字段获取）
AUTH_TOKEN = (
    "Basic eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9"
    ".eyJkYXRhIjoiIiwiZXhwIjoxNzc5MTgwNDg2LCJqdGkiOiJiYTdkYzU2OC0xYmNmLTRmMTgtYTg3Mi1hMTk0ZDVhOTYzOGQiLCJpYXQiOjE3Nzg1NzU2ODYsImlzcyI6IjIwMDQwOSJ9"
    ".m23qy_koNhJzlFuYkhei_0c9CE68PBIGQ0KyF03c8bQ"
)

# Session Cookie（从抓包的 set-cookie 获取，可选）
COOKIES = {
    "acw_tc": "781bad1c17785758105975288eef7e80dce22d67d948e600b3d4a33dedb4a4",
    "JSESSIONID": "5C3C26895D04ACB909D56F471F1A7863",
}

# 请求间隔（秒），避免频率过高
REQUEST_INTERVAL = 2


# ─────────────────────────────────────────
# 请求头（完整复现抓包内容）
# ─────────────────────────────────────────
def build_headers() -> dict:
    return {
        "Accept": "application/json",
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Authentication": AUTH_TOKEN,
        "Connection": "keep-alive",
        "Content-Type": "application/json; charset=UTF-8",
        "Host": "apigateway.gaodun.com",
        "Origin": "https://ocrm.gaodun.com",
        "Referer": "https://ocrm.gaodun.com/",
        "Sec-Ch-Ua": '"Chromium";v="129", "Not=A?Brand";v="8"',
        "Sec-Ch-Ua-Mobile": "?0",
        "Sec-Ch-Ua-Platform": '"Windows"',
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/129.0.6668.101 Safari/537.36"
        ),
        "X-Requested-Extend": json.dumps({"systemName": "SYSTEM-OCRM"}),
        "x_gdsid": "05Y8WO0KX6Osi77halcoyidl5rdqektyfn8nwd0hdmarm",
        "x_gduserid": "97196",
    }


# ─────────────────────────────────────────
# Payload 构造（可批量传入不同数据）
# ─────────────────────────────────────────
def build_payload(
    mobile: str,
    true_name: str = "未留名",
    follow_records: str = "",
    city: str = "石家庄市",
    city_id: int = 19459,
    state: str = "河北省",
    state_id: int = 19458,
    country_id: int = 18784,
    gender: int = 115003,
    intent_project_id: int = 100520929,
    market_channel_id: int = 8218,
    behavior_id: int = 247,
    profession_id: int = 117004,
    rule_type: str = "100521538",
    tags_infos: Optional[list] = None,
    begin_time=None,
    end_time=None,
    contact_type=None,
    specified_date=None,
    line_tel_phone: str = "",
    over_sea_area_code: str = "",
) -> dict:
    """
    构造请求 payload，所有字段均可覆盖。
    最常用的参数（mobile、true_name、follow_records）已提到前面。
    """
    if tags_infos is None:
        tags_infos = [
            {
                "tagCode": 12308048,
                "tagName": "省份",
                "checkStatus": True,
                "typeId": 12308,
                "singleChoice": 1,
            }
        ]

    return {
        "beginTime": begin_time,
        "behaviorId": behavior_id,
        "city": city,
        "cityId": city_id,
        "contactType": contact_type,
        "countryId": country_id,
        "endTime": end_time,
        "followRecords": follow_records,
        "gender": gender,
        "intentProjectId": intent_project_id,
        "lineTelPhone": line_tel_phone,
        "marketChannelId": market_channel_id,
        "mobile": mobile,
        "overSeaAreaCode": over_sea_area_code,
        "professionId": profession_id,
        "ruleType": rule_type,
        "specifiedDate": specified_date,
        "state": state,
        "stateId": state_id,
        "tagsInfos": tags_infos,
        "trueName": true_name,
    }


# ─────────────────────────────────────────
# 核心提交函数
# ─────────────────────────────────────────
def submit_once(session: requests.Session, payload: dict) -> dict:
    """发送单次请求，返回解析后的 JSON 响应。"""
    url = BASE_URL + ENDPOINT
    try:
        resp = session.post(
            url,
            headers=build_headers(),
            json=payload,
            cookies=COOKIES,
            timeout=15,
            verify=True,
        )
        resp.raise_for_status()
        result = resp.json()
        return result
    except requests.exceptions.RequestException as e:
        log.error("请求异常: %s", e)
        return {"error": str(e)}


def is_success(response: dict) -> bool:
    """判断接口是否返回成功（status=0）。"""
    return response.get("status") == 0


# ─────────────────────────────────────────
# 批量提交
# ─────────────────────────────────────────
def batch_submit(records: list[dict], interval: float = REQUEST_INTERVAL):
    """
    批量提交多条记录。
    records: list of kwargs passed to build_payload()
    """
    session = requests.Session()
    success_count = 0
    fail_count = 0

    for i, record in enumerate(records, start=1):
        payload = build_payload(**record)
        log.info("[%d/%d] 提交 mobile=%s name=%s", i, len(records), record.get("mobile"), record.get("true_name", "未留名"))

        response = submit_once(session, payload)

        if is_success(response):
            log.info("  ✅ 成功 | id=%s info=%s", response.get("result", {}).get("id"), response.get("info"))
            success_count += 1
        else:
            log.warning("  ❌ 失败 | 响应: %s", json.dumps(response, ensure_ascii=False))
            fail_count += 1

        if i < len(records):
            time.sleep(interval)

    log.info("─── 批量完成: 成功 %d / 失败 %d ───", success_count, fail_count)
    return success_count, fail_count


# ─────────────────────────────────────────
# 示例入口
# ─────────────────────────────────────────
if __name__ == "__main__":

    # ── 单条提交示例 ──────────────────────
    session = requests.Session()
    single_payload = build_payload(
        mobile="15369397819",
        true_name="未留名",
        follow_records="公务员;知了-石家庄校区;25年毕业，三本机械电子专业求职规划",
    )
    log.info("开始单条提交...")
    resp = submit_once(session, single_payload)
    log.info("响应: %s", json.dumps(resp, ensure_ascii=False, indent=2))

    # ── 批量提交示例 ──────────────────────
    # 取消注释以启用批量模式
    #
    # records = [
    #     {"mobile": "13800000001", "true_name": "张三", "follow_records": "备注A"},
    #     {"mobile": "13900000002", "true_name": "李四", "follow_records": "备注B"},
    #     {"mobile": "15000000003", "true_name": "王五", "follow_records": "备注C"},
    # ]
    # batch_submit(records, interval=2)
