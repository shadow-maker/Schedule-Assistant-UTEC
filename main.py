from scheduleAssistant import ScheduleAssistant

assistant = ScheduleAssistant()
assistant.initWebdriver()
assistant.login(input("Email: "), input("Password: "))
assistant.downloadSchedule()
assistant.pdfToTable()
assistant.tableToDict()
