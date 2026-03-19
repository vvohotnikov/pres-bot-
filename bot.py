# bot.py
import os, asyncio, tempfile
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import CommandStart
from openai import AsyncOpenAI
from pydub import AudioSegment
from dotenv import load_dotenv
import subprocess

load_dotenv()

bot = Bot(token=os.getenv("TELEGRAM_TOKEN"))
dp = Dispatcher()
client = AsyncOpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    base_url="https://api.proxyapi.ru/openai/v1"
)

# --- СИСТЕМНЫЙ ПРОМТ ---
SYSTEM_PROMPT = """
Ты — ассистент для создания презентаций. 
Тебе дают расшифровку голосового сообщения. 
Твоя задача — структурировать его в Markdown-презентацию.

Правила:
- Используй ТОЛЬКО то, что сказал пользователь. Не додумывай ничего лишнего.
- Каждый слайд: # Заголовок слайда + 3-5 коротких буллетов
- Максимум 10 слайдов
- Первый слайд — титульный с темой и подзаголовком
- Последний слайд — ключевые выводы

Формат вывода — чистый Markdown, ничего лишнего.
"""

# --- ХЭНДЛЕРЫ ---
@dp.message(CommandStart())
async def start(message: Message):
    await message.answer(
        "Привет! 🎤\n\n"
        "Наговори голосовое — я соберу из него презентацию.\n"
        "Можешь отправить несколько войсов подряд, затем напиши /build"
    )

# Хранилище транскриптов (в памяти, для прода — Redis/БД)
transcripts = {}

@dp.message(F.voice)
async def handle_voice(message: Message):
    user_id = message.from_user.id
    await message.answer("⏳ Транскрибирую...")

    # Скачиваем .oga файл
    voice_file = await bot.get_file(message.voice.file_id)
    with tempfile.TemporaryDirectory() as tmp:
        oga_path = f"{tmp}/voice.oga"
        mp3_path = f"{tmp}/voice.mp3"
        await bot.download_file(voice_file.file_path, oga_path)

        # Конвертируем в mp3
        AudioSegment.from_ogg(oga_path).export(mp3_path, format="mp3")

        # Whisper транскрипция
        with open(mp3_path, "rb") as f:
            result = await client.audio.transcriptions.create(
                model="whisper-1", file=f, language="ru"
            )

    text = result.text
    # Накапливаем транскрипты от пользователя
    transcripts[user_id] = transcripts.get(user_id, "") + "\n\n" + text
    await message.answer(f"✅ Записал:\n\n_{text}_\n\nОтправь ещё войс или /build", parse_mode="Markdown")

@dp.message(F.text == "/build")
async def build_presentation(message: Message):
    user_id = message.from_user.id
    if user_id not in transcripts or not transcripts[user_id].strip():
        await message.answer("Сначала отправь голосовое сообщение 🎤")
        return

    await message.answer("🔨 Собираю презентацию...")

    # GPT-4o структурирует транскрипт в Markdown
    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": transcripts[user_id]}
        ]
    )
    md_content = response.choices[0].message.content

    # Конвертируем Markdown → PPTX через Pandoc с вашим шаблоном
    with tempfile.TemporaryDirectory() as tmp:
        md_path = f"{tmp}/presentation.md"
        pptx_path = f"{tmp}/presentation.pptx"

        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)

        subprocess.run([
            "pandoc", md_path,
            "-o", pptx_path,
            "--reference-doc=template.pptx"  # 👈 ваш фирменный шаблон
        ], check=True)

        await message.answer_document(
            FSInputFile(pptx_path, filename="presentation.pptx"),
            caption="🎉 Готово! Презентация на основе твоих слов."
        )

    # Сбрасываем буфер
    del transcripts[user_id]

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
