def emailMsg(duplicate_count, include_count, download_failed_count, totalRun, endTime, startTime):

    summary = (
        "<strong>TXPUC0 DOCUMENT COLLECTOR COMPLETE.</strong><br><br>"
        "Attached log file of the run.<br><br>"
        f"Total run time: {totalRun(endTime, startTime)}<br>"
        f"Duplicates: {duplicate_count}<br>"
        f"Includes: {include_count}<br>"
        f"Error to Download: {download_failed_count}<br><br>"
        "<strong>-----------------------------------------------------------</strong><br><br>"
    )
    details = "<br><br>Thanks!<br>TXPUC0 Document Collector - v4.0"

    emailBody = summary + details

    return emailBody