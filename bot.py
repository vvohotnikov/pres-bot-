# bot.py
import os, asyncio, tempfile, copy
from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, FSInputFile
from aiogram.filters import CommandStart
from openai import AsyncOpenAI
from pydub import AudioSegment
from dotenv import load_dotenv
from pptx import Presentation
from pptx.util import Pt
import json

load_dotenv()

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pptx")

bot = Bot(token=os.getenv("TELEGRAM_TOKEN"))
dp = Dispatcher()
client = AsyncOpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
    base_url="https://api.proxyapi.ru/openai/v1"
)  # ← скобка была не закрыта

TITLE_SLIDE_IDX   = 0
CONTENT_SLIDE_IDX = 8

TITLE_SHAPES = {
    "title":    "Google Shape;227;p29",
    "subtitle": "Google Shape;228;p29",
    "speaker":  "Google Shape;229;p29",
    "role":     "Google Shape;231;p29",
    "speaker2": "Google Shape;230;p29",
    "role2":    "Google Shape;232;p29",
}  # ← скобка была не закрыта

CONTENT_SHAPES = {
    "title":    "Google Shape;436;p42",
    "subtitle": "Google Shape;437;p42",
    "bullet0":  "Google Shape;438;p42",
    "bullet1":  "Google Shape;439;p42",
    "bullet2":  "Google Shape;440;p42",
    "bullet3":  "Google Shape;441;p42",
}  # ← скобка была не закрыта

SYSTEM_PROMPT = """
Ты — ассистент для создания презентаций на русском языке.
Тебе дают расшифровку голосового сообщения.

Верни ТОЛЬКО валидный JSON без markdown-блоков, по следующей структуре:
{
  "title": "Заголовок презентации",
  "subtitle": "Одна фраза о чём презентация",
  "speaker": "Имя спикера если упомянуто, иначе пустая строка",
  "role": "Должность если упомянута, иначе пустая строка",
  "slides": [
    {
      "title": "Заголовок слайда",
      "subtitle": "Короткий подзаголовок или тема слайда",
      "bullets": [
        "Первый тезис — полное предложение 10-15 слов",
        "Второй тезис — полное предложение 10-15 слов",
        "Третий тезис — полное предложение 10-15 слов",
        "Четвёртый тезис — полное предложение 10-15 слов"
      ]
    }
  ]
}

Правила:
- Используй ВСЁ что сказал пользователь, раскрывай каждую мысль полностью
- От 4 до 8 слайдов в "slides"
- Ровно 4 буллета на каждый слайд
- Последний слайд всегда "Ключевые выводы"
- Только JSON, никаких пояснений
"""

transcripts = {}


def set_shape_text(slide, shape_name: str, text: str):
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_text_frame:
            tf = shape.text_frame
            for para in tf.paragraphs:
                if para.runs:
                    para.runs[0].text = text
                    for run in para.runs[1:]:
                        run.text = ''
                    return True
    return False


def duplicate_slide(prs: Presentation, slide_idx: int):
    template_slide = prs.slides[slide_idx]
    layout = template_slide.slide_layout
    new_slide = prs.slides.add_slide(layout)
    sp_tree = new_slide.shapes._spTree
    for sp in list(sp_tree)[2:]:
        sp_tree.remove(sp)
    for shape in template_slide.shapes:
        sp_tree.append(copy.deepcopy(shape._element))
    return new_slide


def build_pptx( dict, output_path: str):  # ← было: def build_pptx( dict, ...
    prs = Presentation(TEMPLATE_PATH)
    slides_data = data.get('slides', [])

    new_title_slide = duplicate_slide(prs, TITLE_SLIDE_IDX)
    new_content_slides = []
    for _ in slides_  # ← было: for _ in slides_
        new_content_slides.append(duplicate_slide(prs, CONTENT_SLIDE_IDX))

    set_shape_text(new_title_slide, TITLE_SHAPES["title"],    data.get('title', ''))
    set_shape_text(new_title_slide, TITLE_SHAPES["subtitle"], data.get('subtitle', ''))
    set_shape_text(new_title_slide, TITLE_SHAPES["speaker"],  data.get('speaker', ''))
    set_shape_text(new_title_slide, TITLE_SHAPES["role"],     data.get('role', ''))
    set_shape_text(new_title_slide, TITLE_SHAPES["speaker2"], '')
    set_shape_text(new_title_slide, TITLE_SHAPES["role2"],    '')

    for i, slide_info in enumerate(slides_data):
        s = new_content_slides[i]
        bullets = slide_info.get('bullets', [])
        set_shape_text(s, CONTENT_SHAPES["title"],    slide_info.get('title', ''))
        set_shape_text(s, CONTENT_SHAPES["subtitle"], slide_info.get('subtitle', ''))
        set_shape_text(s, CONTENT_SHAPES["bullet0"],  bullets[0] if len(bullets) > 0 else '')
        set_shape_text(s, CONTENT_SHAPES["bullet1"],  bullets[1] if len(bullets) > 1 else '')
        set_shape_text(s, CONTENT_SHAPES["bullet2"],  bullets[2] if len(bullets) > 2 else '')
        set_shape_text(s, CONTENT_SHAPES["bullet3"],  bullets[3] if len(bullets) > 3 else '')

    sldIdLst = prs.slides._sldIdLst
    all_ids = list(sldIdLst)
    for sldId in all_ids[:10]:
        sldIdLst.remove(sldId)

    prs.save(output_path)


@dp.message(CommandStart())
async def start(message: Message):
    await message.answer(
        "Привет! 🎤\n\n"
        "Наговори голосовое — я соберу из него презентацию.\n"
        "Можешь отправить несколько войсов подряд, затем напиши /build"
    )  # ← скобка была не закрыта


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
            )  # ← скобка была не закрыта

    text = result.text
    transcripts[user_id] = transcripts.get(user_id, "") + "\n\n" + text
    await message.answer(
        f"✅ Записал:\n\n_{text}_\n\nОтправь ещё войс или /build",
        parse_mode="Markdown"
    )  # ← скобка была не закрыта


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
    )  # ← скобка была не закрыта

    raw = response.choices[0].message.content.strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        await message.answer("❌ GPT вернул некорректный JSON. Попробуй ещё раз.")
        return

    with tempfile.TemporaryDirectory() as tmp:
        pptx_path = f"{tmp}/presentation.pptx"
        build_pptx(data, pptx_path)
        await message.answer_document(
            FSInputFile(pptx_path, filename="presentation.pptx"),
            caption="🎉 Готово! Презентация на основе твоих слов."
        )  # ← скобка была не закрыта

    del transcripts[user_id]


async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
