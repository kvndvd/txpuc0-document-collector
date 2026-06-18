import logging, time

def totalRun(endTime, startTime):
    duration = int(endTime - startTime)
    hours, remainder = divmod(duration, 3600)
    minutes, seconds = divmod(remainder, 60)

    parts = []
    if hours > 0:
        parts.append(f"{hours} hr" + ("s" if hours != 1 else ""))
    if minutes > 0:
        parts.append(f"{minutes} min" + ("s" if minutes != 1 else ""))
    if seconds > 0 or not parts:
        parts.append(f"{seconds} sec" + ("s" if seconds != 1 else ""))

    return " ".join(parts)