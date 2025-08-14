from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from openpyxl import load_workbook

async def find(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if update.message and update.message.text:
        workbook = load_workbook('test.xlsx')
        sheet = workbook.active
        response_message = ''

        for row in sheet.iter_rows():
            parsed_message = update.message.text.replace("/find", "").strip() 
            if parsed_message == str(row[0].value):
                response_message += row[1].value
                
        await update.message.reply_text(f"{response_message}")

def main():
    app = ApplicationBuilder().token("BOT_TOKEN").build()
    app.add_handler(CommandHandler("find", find))
    app.run_polling()

if __name__ == "__main__":
    main()
