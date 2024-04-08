import asyncio
import logging
from aiogram import Bot
from aiogram import Dispatcher,types
from aiogram.filters import  CommandStart,Command


TOKEN = '6887247322:AAFepKNzXiaRn0myaN_U-aabZOGnD9Ulc5g'

bot = Bot(token=TOKEN)
dp = Dispatcher()

@dp.message(CommandStart())
async def handle_start(message: types.Message):
    await message.answer(text=f'Hello, {message.from_user.full_name}!')

@dp.message(Command("help"))
async def handle_help(message: types.Message):
    await message.answer(text='Ya prosto echo bot')

@dp.message()
async def echo_message(message: types.Message):
    await bot.send_message(
        chat_id=message.chat.id,
        text='Start processing')

    await bot.send_message(
        chat_id=message.chat.id,
        text='see message',
        reply_to_message_id=message.message_id,
    )
    await message.answer(text='Wait')
    try:
        await message.send_copy(chat_id=message.chat.id)
    except TypeError:
        await  message.reply(text = 'dont know')

    #if message.text :
    #    await  message.reply(
    #        text = message.text
    #    )
    #elif message.sticker:
    #    await message.reply_sticker(sticker=message.sticker.file_id)
    #elif message.photo:
    #    await message.reply_photo(photo=message.photo.file_id)
    #else:
    #    await  message.reply(text = 'dont know')

async def main():
    logging.basicConfig(level=logging.INFO)
    await dp.start_polling(bot)


if __name__=='__main__':
    asyncio.run(main())