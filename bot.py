# bot.py
import os, asyncio, tempfile
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import CommandStart
from openai import AsyncOpenAI
from pydub import AudioSegment
from dotenv import load_dotenv
from pptx import Presentation
from pptx.util import Inches, Pt

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

# Хранилище транскриптов (в памяти)
transcripts = {}

@dp.message(F.voice)
async def handle_voice(message: Message):
    user_id = message.from_user.id
    await message.answer("⏳ Транскрибирую...")

    voice_file = await bot.get_file(message.voice.file_id)
    with tempfile.TemporaryDirectory() as tmp:
        oga_path = f"{tmp}/voice.oga"
        mp3_path = f"{tmp}/voice.mp3"
        await bot.download_file(voice_file.file_path, oga_path)

        AudioSegment.from_ogg(oga_path).export(mp3_path, format="mp3")

        with open(mp3_path, "rb") as f:
            result = await client.audio.transcriptions.create(
                model="whisper-1", file=f, language="ru"
            )

    text = result.text
    transcripts[user_id] = transcripts.get(user_id, "") + "\n\n" + text
    await message.answer(f"✅ Записал:\n\n_{text}_\n\nОтправь ещё войс или /build", parse_mode="Markdown")


def build_pptx(md_content: str, output_path: str):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    current_title = None
    current_bullets = []

    def flush_slide():
        if current_title is None:
            return
        if not current_bullets:
            layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = current_title
        else:
            layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = current_title
            tf = slide.placeholders[1].text_frame
            tf.text = current_bullets[0]
            for b in current_bullets[1:]:
                p = tf.add_paragraph()
                p.text = b

    for line in md_content.splitlines():
        line = line.strip()
        if line.startswith("# "):
            flush_slide()
            current_title = line[2:].strip()
            current_bullets = []
        elif line.startswith("## "):
            flush_slide()
            current_title = line[3:].strip()
            current_bullets = []
        elif line.startswith(("- ", "* ", "• ")):
            current_bullets.append(line[2:].strip())

    flush_slide()
    prs.save(output_path)


@dp.message(F.text == "/build")
async def build_presentation(message: Message):
    user_id = message.from_user.id
    if user_id not in transcripts or not transcripts[user_id].strip():
        await message.answer("Сначала отправь голосовое сообщение 🎤")
        return

    await message.answer("🔨 Собираю презентацию...")

    response = await client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": transcripts[user_id]}
        ]
    )
    md_content = response.choices[0].message.content

    with tempfile.TemporaryDirectory() as tmp:
        pptx_path = f"{tmp}/presentation.pptx"
        build_pptx(md_content, pptx_path)

        await message.answer_document(
            FSInputFile(pptx_path, filename="presentation.pptx"),
            caption="🎉 Готово! Презентация на основе твоих слов."
        )

    del transcripts[user_id]


async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
