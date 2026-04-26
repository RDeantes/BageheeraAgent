from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, ContextTypes, filters
import sys
import os

from BagheeraBrain import procesar

print("BOT INICIADO")
print("HOLA JEFA COMO PUEDO AYUDARTE HOY")


# 🔑 TU TOKEN (pon el tuyo aquí)
TOKEN = "8614896779:AAEQTEffYyNIBN7OBnTij8qFfexqy_02aLQ"


# =========================================================
# 🎯 RESPONDER MENSAJES
# =======================================================
async def responder(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mensaje_usuario = update.message.text.upper()

    print("MENSAJE RECIBIDO:", mensaje_usuario)

    # 🔴 Si presiona ESC → rompe el ciclo y reinicia
    if mensaje_usuario == "ESC":
        await update.message.reply_text("🔄 Reiniciando el proceso...")
        print("Reiniciando...")
        os.execv(sys.executable, ['python'] + sys.argv)

    resultado = procesar(mensaje_usuario)

    # 🟢 Si es PDF → lo envía
    if isinstance(resultado, str) and resultado.endswith(".pdf"):
        try:
            with open(resultado, "rb") as f:
                await context.bot.send_document(
                    chat_id=update.effective_chat.id,
                    document=f
                )

            await update.message.reply_text("📄 Contrato generado y enviado correctamente")

        except Exception as e:
            print("ERROR AL ENVIAR PDF:", e)

            # 🧠 Manejo inteligente de timeout
            if "Timed out" in str(e):
                await update.message.reply_text(
                    "📄 El contrato probablemente ya fue enviado (conexión lenta)"
                )
            else:
                await update.message.reply_text("❌ Error real al enviar el PDF")

    # 🟡 Si es texto → responde normal
    elif resultado:
        await update.message.reply_text(resultado)

    # 🔴 Si algo falló
    else:
        await update.message.reply_text("⚠️ Ocurrió un error")


# =========================================================
# 🚀 INICIAR BOT
# =========================================================
if __name__ == "__main__":
    app = ApplicationBuilder().token(TOKEN).build()

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, responder))

    print("Bagheera está corriendo...")

    app.run_polling()