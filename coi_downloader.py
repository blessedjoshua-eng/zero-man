"""
Zero Man — COI PDF Bulk Downloader
Backend logic, imported and called by app.py

Changes from original script
─────────────────────────────
  • No hardcoded paths or globals
  • process_customers() accepts all config as parameters
  • column_mapping  : renames Excel columns to required field names
  • All print() replaced with log_cb(dict) for live SSE streaming
  • send_email_with_pdf() takes smtp_cfg dict
  • PDF saved as  {number_value}_COI.pdf
  • API / encryption logic is identical to the original
"""

import os
import io
import time
import json
import base64
import hashlib
import smtplib
import pandas as pd
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import requests
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad


# ─────────────────────────────────────────────
#  API Constants
# ─────────────────────────────────────────────
BASE_URL         = "https://www.indiafirstlife.com"
TOKEN_URL        = f"{BASE_URL}/content/ifliwebsite/in/osgi_token.json"
TOKEN_BODY_STEP1 = "siteidentifier=api_coi_tokenization"
TOKEN_BODY_STEP2 = "siteidentifier=WebsiteToken"
STEP1_URL        = "https://apig.indiafirstlife.com/getgroupcoidetails/"
STEP2_URL        = "https://apig.indiafirstlife.com/website_data/getFciWebhookDetails"


# ─────────────────────────────────────────────
#  Encryption Constants
# ─────────────────────────────────────────────
SIMPLE_KEY     = b"tokentokentokentokentokentokenwe"
SIMPLE_IV      = b"encryptionIntVec"
CODE_KEY       = "d6163f0659cfe4196dc03c2c29aab06f10cb0a79cdfc74a45da2d72358712e80"
PBKDF2_KEY_LEN = 32
PBKDF2_ITERS   = 100


# ─────────────────────────────────────────────
#  Product → Trigger Name Map
# ─────────────────────────────────────────────
TRIGGER_MAP = {
    "GLP"      : "IFL_OnDemand_GLP",
    "HEALTH"   : "IFL_OnDemand_Hospicare",
    "MICRO"    : "IFL_OnDemand_Micro",
    "GTL"      : "IFL_OnDemand_GTL",
    "GCL"      : "IFL_OnDemand_GroupCreditLife",
    "GCL PLUS" : "IFL_OnDemand_GCLPlus",
}


# ─────────────────────────────────────────────
#  Request Headers
# ─────────────────────────────────────────────
HEADERS_BASE = {
    "User-Agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept"     : "application/json, text/plain, */*",
    "Origin"     : "https://www.indiafirstlife.com",
    "Referer"    : "https://www.indiafirstlife.com/common-coi",
}


# ─────────────────────────────────────────────
#  Encryption Helpers
# ─────────────────────────────────────────────
def simple_aes_encrypt(plaintext: str) -> str:
    cipher = AES.new(SIMPLE_KEY, AES.MODE_CBC, SIMPLE_IV)
    ct = cipher.encrypt(pad(plaintext.encode("utf-8"), AES.block_size))
    return base64.b64encode(ct).decode("utf-8")


def simple_aes_decrypt(ciphertext_b64: str) -> str:
    ct = base64.b64decode(ciphertext_b64)
    cipher = AES.new(SIMPLE_KEY, AES.MODE_CBC, SIMPLE_IV)
    return unpad(cipher.decrypt(ct), AES.block_size).decode("utf-8")


def _pbkdf2_sha1(password: bytes, salt: bytes, iterations: int, key_len: int) -> bytes:
    return hashlib.pbkdf2_hmac("sha1", password, salt, iterations, dklen=key_len)


def pbkdf2_aes_encrypt(passphrase: str, plaintext: str) -> str:
    salt   = os.urandom(32)
    iv     = os.urandom(16)
    key    = _pbkdf2_sha1(passphrase.encode("utf-8"), salt, PBKDF2_ITERS, PBKDF2_KEY_LEN)
    cipher = AES.new(key, AES.MODE_CBC, iv)
    ct     = cipher.encrypt(pad(plaintext.encode("utf-8"), AES.block_size))
    return salt.hex() + iv.hex() + base64.b64encode(ct).decode("utf-8")


def pbkdf2_aes_decrypt(passphrase: str, ciphertext: str) -> str:
    salt   = bytes.fromhex(ciphertext[:64])
    iv     = bytes.fromhex(ciphertext[64:96])
    ct     = base64.b64decode(ciphertext[96:])
    key    = _pbkdf2_sha1(passphrase.encode("utf-8"), salt, PBKDF2_ITERS, PBKDF2_KEY_LEN)
    cipher = AES.new(key, AES.MODE_CBC, iv)
    return unpad(cipher.decrypt(ct), AES.block_size).decode("utf-8")


# ─────────────────────────────────────────────
#  Token Helper
# ─────────────────────────────────────────────
def get_bearer_token(session: requests.Session, token_body: str) -> str:
    resp = session.post(
        TOKEN_URL,
        data=token_body,
        headers={**HEADERS_BASE, "Content-Type": "application/x-www-form-urlencoded"},
        timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


# ─────────────────────────────────────────────
#  Step 1 — Fetch Policy Details
# ─────────────────────────────────────────────
def get_policy_details(session, token, product_type, number_type, number_value, master_policy=""):
    payload = {
        "Product_Type"        : product_type,
        "Master_Policy_Number": master_policy,
        "Filter_Type"         : number_type,
        "Filter_Value"        : number_value,
    }
    encrypted = simple_aes_encrypt(json.dumps(payload))
    resp = session.post(
        STEP1_URL,
        json={"data": encrypted},
        headers={**HEADERS_BASE, "Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=30,
    )
    resp.raise_for_status()
    resp_json = resp.json()
    if "data" not in resp_json:
        raise ValueError(f"Unexpected Step 1 response: {resp_json}")
    decrypted = json.loads(simple_aes_decrypt(resp_json["data"]))
    if isinstance(decrypted, dict) and "getGroupCoiResponse" in decrypted:
        records = decrypted["getGroupCoiResponse"].get("data", [])
        if not records:
            raise ValueError(f"No policy records found for {number_value}")
        return records[0]
    return decrypted


# ─────────────────────────────────────────────
#  Step 2 — Download COI PDF
# ─────────────────────────────────────────────
def download_coi_pdf(session, token, trigger_name, rec, dob, financial_year, number_type, number_value):
    dob_normalised = dob.replace("-", "/")

    webhook_payload = {
        "triggerName"        : trigger_name,
        "to"                 : [], "cc": [], "bcc": [],
        "mobileNumber"       : rec.get("mobileNo", ""),
        "communicationCode"  : "0",
        "attachementRequired": "yes",
        "isPdfRequired"      : True,
        "data": {
            "HeaderName"             : "",
            "UIN"                    : rec.get("UIN", ""),
            "MobileNo"               : rec.get("mobileNo", ""),
            "MasterPolicyholderName" : rec.get("masterPolicyHolderName", ""),
            "MasterPolicyNo"         : rec.get("masterPolicyNumber", ""),
            "BaseSumAssured"         : rec.get("baseSumAssured", ""),
            "premiumPaid"            : rec.get("dprem01", ""),
            "PremiumPayingFrequency" : rec.get("premiumPayingFrequency", ""),
            "COINumber"              : rec.get("coiNo", ""),
            "MemberNumber"           : rec.get("memberNumber", ""),
            "MemberAge"              : rec.get("memberAge", ""),
            "JointMemberName"        : rec.get("jointMember", ""),
            "jointborrowerDOB"       : rec.get("jointMemberDOB", ""),
            "RelationshipOfNominee"  : rec.get("relationshipOfNominee", ""),
            "JointLifeSumAssured"    : rec.get("jointLifeSumAssured", ""),
            "typeOfCover"            : rec.get("typeOfCover", ""),
            "loanNo"                 : rec.get("loanNo", ""),
            "borrowerName"           : rec.get("borrowerName", ""),
            "borrowerDOB"            : rec.get("borrowerDOB", ""),
            "borrowerGender"         : rec.get("borrowerGender", ""),
            "SumAssured"             : rec.get("sumassured", ""),
            "totalsumassured"        : rec.get("totalsumassured", ""),
            "PLANNO"                 : rec.get("benefitOption", ""),
            "PAYMTH"                 : rec.get("paymth", ""),
            "coverTerm"              : rec.get("coverTerm", ""),
            "premiumPaymentTerm"     : rec.get("premiumPaymentTerm", ""),
            "Pay_Frequency"          : rec.get("premiumPayingFrequency", ""),
            "coverCommencementDate"  : rec.get("coverCommencementDate", ""),
            "coverEndDate"           : rec.get("coverEndDate", ""),
            "DOB"                    : dob_normalised,
            "Financial_Year"         : financial_year,
            "Filter_Type"            : number_type,
            "Filter_Value"           : number_value,
            "Product_Type"           : trigger_name.split("_")[-1] if "_" in trigger_name else "",
        },
    }

    encrypted = pbkdf2_aes_encrypt(CODE_KEY, json.dumps(webhook_payload))
    resp = session.post(
        STEP2_URL,
        json={"data": encrypted},
        headers={**HEADERS_BASE, "Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=60,
    )
    resp.raise_for_status()
    resp_json = resp.json()

    if resp_json.get("header", {}).get("statusCode") != "200":
        raise ValueError(f"API error: {resp_json.get('header', {}).get('responseMessage')}")

    decrypted = pbkdf2_aes_decrypt(CODE_KEY, resp_json["response"])
    resp_data = json.loads(decrypted)

    if resp_data.get("status") != "200":
        raise ValueError(f"PDF fetch failed: {resp_data}")

    return base64.b64decode(resp_data["fileContent"])


# ─────────────────────────────────────────────
#  Email Helper
# ─────────────────────────────────────────────
def send_email_with_pdf(smtp_cfg: dict, to_email: str, customer_name: str,
                        pdf_bytes: bytes, pdf_filename: str):
    msg            = MIMEMultipart()
    msg["From"]    = smtp_cfg["user"]
    msg["To"]      = to_email
    msg["Subject"] = "Your Certificate of Insurance (COI) — IndiaFirst Life"
    greeting = f"Dear {customer_name}," if customer_name else "Dear Customer,"
    body = (
        f"{greeting}\n\n"
        "Please find your Certificate of Insurance (COI) attached to this email.\n\n"
        "If you have any queries, please contact IndiaFirst Life customer care.\n\n"
        "Regards,\nIndiaFirst Life Team"
    )
    msg.attach(MIMEText(body, "plain"))
    part = MIMEBase("application", "pdf")
    part.set_payload(pdf_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{pdf_filename}"')
    msg.attach(part)
    with smtplib.SMTP(smtp_cfg["host"], smtp_cfg["port"]) as srv:
        srv.starttls()
        srv.login(smtp_cfg["user"], smtp_cfg["password"])
        srv.send_message(msg)


# ─────────────────────────────────────────────
#  Main Processor
# ─────────────────────────────────────────────
def process_customers(excel_bytes: bytes, output_dir: str,
                      smtp_cfg: dict, send_email:
