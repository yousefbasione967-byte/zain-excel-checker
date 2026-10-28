import os
import io
import logging
from decimal import Decimal, InvalidOperation

import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, ContextTypes, filters

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
# Optional: set ALLOWED_USER_IDS to comma-separated Telegram numeric user IDs.
ALLOWED_USER_IDS = {
    x.strip() for x in os.environ.get("ALLOWED_USER_IDS", "").split(",") if x.strip()
}

def is_allowed(update: Update) -> bool:
    if not ALLOWED_USER_IDS:
        return True
    user = update.effective_user
    return bool(user and str(user.id) in ALLOWED_USER_IDS)

def normalize_account(value):
    if pd.isna(value):
        return ""
    # Keep account IDs as text; Excel may store them as numbers.
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    # Convert Arabic-Indic digits to Western digits.
    trans = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    return s.translate(trans)

def parse_amount(value):
    if pd.isna(value):
        return None
    s = str(value).strip().replace(",", "").replace("٬", "").replace("٫", ".")
    trans = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    s = s.translate(trans).replace("ر.س", "").replace("SAR", "").strip()
    try:
        amount = Decimal(s)
        if not amount.is_finite() or amount < 0:
            return None
        return float(amount)
    except (InvalidOperation, ValueError):
        return None

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("غير مصرح لك باستخدام هذا البوت.")
        return
    await update.message.reply_text(
        "أهلًا! ابعت ملف Excel بصيغة .xlsx يحتوي على عمودين:\n"
        "account = رقم الحساب (10 أرقام)\n"
        "original_amount = المبلغ الأصلي\n\n"
        "ملاحظة: هذه نسخة أولية لمعالجة الشيت فقط؛ لم يتم ربطها بنظام زين بعد، "
        "ولن تخمّن حالة السداد."
    )

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("غير مصرح لك باستخدام هذا البوت.")
        return

    doc = update.message.document
    if not doc or not doc.file_name.lower().endswith(".xlsx"):
        await update.message.reply_text("من فضلك ابعت ملف Excel بصيغة .xlsx.")
        return

    if doc.file_size and doc.file_size > 10 * 1024 * 1024:
        await update.message.reply_text("الملف أكبر من 10 ميجابايت. قسّمه لملفات أصغر.")
        return

    await update.message.reply_text("استلمت الملف. بفحص الأعمدة وأجهز تقريرًا أوليًا...")
    tg_file = await context.bot.get_file(doc.file_id)
    raw = await tg_file.download_as_bytearray()

    try:
        df = pd.read_excel(io.BytesIO(raw), dtype={"account": str})
    except Exception as exc:
        logger.exception("Could not read workbook: %s", exc)
        await update.message.reply_text("مش قادر أقرأ الملف. تأكد إنه ملف .xlsx سليم.")
        return

    # Accept common Arabic alternatives for column headings.
    aliases = {
        "account": ["account", "account_number", "رقم الحساب", "رقم حساب", "الحساب"],
        "original_amount": ["original_amount", "amount", "المبلغ", "المبلغ الأصلي", "المبلغ المستحق"],
    }
    cols = {str(c).strip().lower(): c for c in df.columns}
    def find_col(key):
        for alias in aliases[key]:
            if alias.lower() in cols:
                return cols[alias.lower()]
        return None

    account_col = find_col("account")
    amount_col = find_col("original_amount")
    if not account_col or not amount_col:
        await update.message.reply_text(
            "مش لاقي الأعمدة المطلوبة.\n"
            "سمّي العمودين account و original_amount، أو استخدم: رقم الحساب والمبلغ المستحق."
        )
        return

    output_rows = []
    for _, row in df.iterrows():
        account = normalize_account(row[account_col])
        amount = parse_amount(row[amount_col])
        valid_account = len(account) == 10 and account.isdigit()
        if not account and amount is None:
            continue
        if not valid_account:
            status = "رقم حساب غير صالح"
        elif amount is None:
            status = "مبلغ غير صالح"
        else:
            status = "بانتظار ربط مصدر زين"
        output_rows.append({
            "account": account,
            "original_amount": amount,
            "current_amount": None,
            "difference": None,
            "status": status,
        })

    if not output_rows:
        await update.message.reply_text("الملف مفيهوش صفوف صالحة للمعالجة.")
        return

    out = io.BytesIO()
    result = pd.DataFrame(output_rows, columns=[
        "account", "original_amount", "current_amount", "difference", "status"
    ])
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Results")
    out.seek(0)
    out.name = "zain_check_template.xlsx"

    valid_count = sum(1 for r in output_rows if len(r["account"]) == 10 and r["account"].isdigit())
    await update.message.reply_document(
        document=out,
        filename="zain_check_template.xlsx",
        caption=(
            f"تم تجهيز التقرير الأولي.\n"
            f"عدد الصفوف: {len(output_rows)}\n"
            f"أرقام الحسابات الصحيحة شكليًا: {valid_count}\n\n"
            "تنبيه: المبلغ الحالي وحالة السداد غير متحقق منهما بعد؛ "
            "يلزم ربط مصدر زين معتمد قبل استخدام التقرير لاتخاذ قرارات."
        ),
    )

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update):
        await update.message.reply_text("غير مصرح لك باستخدام هذا البوت.")
        return
    await update.message.reply_text(
        "ارفع ملف .xlsx فيه عمود account وعمود original_amount. "
        "هذه النسخة تفحص تنسيق البيانات فقط ولا تستعلم من زين حتى الآن."
    )

def main():
    if not BOT_TOKEN:
        raise RuntimeError("Set BOT_TOKEN as an environment variable.")
    app = Application.builder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
