from . import bot
import sys

if len(sys.argv) >= 5 and str(sys.argv[1]).lower() == "--execution".lower():
    bot.Bot.action(None)
else:
    bot.Bot.main()
