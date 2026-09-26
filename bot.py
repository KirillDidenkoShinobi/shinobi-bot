        f"🔗 Создано ссылок: "
        f"<b>{total_links}</b>\n"
        f"👥 Приглашено: "
        f"<b>{total_invited}</b>",
        parse_mode="HTML"
    )


# =========================================================
# EXCEL
# =========================================================

@dp.callback_query(
    F.data == "export_excel"
)
async def export_excel_callback(
    callback: CallbackQuery
):

    if not is_admin(
        callback.from_user.id
    ):

        await callback.answer(
            "❌ Нет доступа",
            show_alert=True
        )

        return

    await callback.answer(
        "📊 Создаю Excel..."
    )

    try:

        filename = create_excel()

        document = FSInputFile(
            filename
        )

        total_invited = sum(
            user.get(
                "total_invited",
                0
            )
            for user in data[
                "users"
            ].values()
        )

        await callback.message.answer_document(
            document=document,
            caption=(
                "📊 <b>SHINOBI TEAM</b>\n"
                "<b>Реферальный отчёт</b>\n\n"
                f"👤 Пользователей: "
                f"<b>{len(data['users'])}</b>\n"
                f"🔗 Ссылок: "
                f"<b>{len(data['links'])}</b>\n"
                f"👥 Приглашено: "
                f"<b>{total_invited}</b>\n\n"
                "📄 Лист 1 — ссылки\n"
                "📄 Лист 2 — приглашённые"
            ),
            parse_mode="HTML"
        )

    except Exception as error:

        await callback.message.answer(
            "❌ Ошибка создания Excel:\n\n"
            f"<code>{error}</code>",
            parse_mode="HTML"
        )

    finally:

        if os.path.exists(
            EXCEL_FILE
        ):

            try:

                os.remove(
                    EXCEL_FILE
                )

            except OSError:
                pass


# =========================================================
# ЗАПУСК
# =========================================================

async def main():
    print("Удаляю старый webhook...")

    # Удаляем webhook перед запуском polling
    await bot.delete_webhook(drop_pending_updates=True)

    print("Webhook удалён.")
    print("=" * 50)
    print(" SHINOBI TEAM REFERRAL BOT")
    print(" БОТ ЗАПУЩЕН")
    print("=" * 50)

    await dp.start_polling(
        bot,
        allowed_updates=[
            "message",
            "callback_query",
            "chat_member",
            "my_chat_member",
            "chat_join_request"
        ]
    )


if __name__ == "__main__":
    asyncio.run(main())
