import script
import time
import schedule

schedule.every().hour.at(":30").do(script)
schedule.every().hour.at(":00").do(script)

while True:
    schedule.run_pending()
    time.sleep(1)