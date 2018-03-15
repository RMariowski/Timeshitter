#!/usr/bin/env python

from operations import seconds_to_full_hours, process_csv, process_docx

FILE_NAME = 'worklog.csv'
DOCX_FILE_NAME = 'rozliczenie.docx'
HEADER_DATE = 'Start Time'
HEADER_SPENT = 'Time Spent (s)'

def main():
    with open(FILE_NAME, encoding="utf8") as csvfile:
        total_time = 0  # [SECONDS]

        worklogs = process_csv(csvfile, HEADER_DATE, HEADER_SPENT)

        for key, value in worklogs.items():
            hours_spent = seconds_to_full_hours(worklogs[key])
            worklogs[key] = hours_spent
            total_time += hours_spent
            print(key, value)

        print('Total time: ' + str(total_time))

        process_docx(DOCX_FILE_NAME, worklogs, total_time)


if __name__ == "__main__":
    main()
