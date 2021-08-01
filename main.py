from scheduleAssistant import ScheduleAssistant

assistant = ScheduleAssistant(input("Email: "), input("Password: "))
assistant.initWebdriver()
assistant.login()
assistant.downloadScheduleData()
assistant.pdfToTable()
assistant.tableToDict()
