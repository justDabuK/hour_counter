import xlrd
import os
import sys

# global definitions
months = ["Januar", "Februar", "Maerz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober",
          "November", "Dezember"]
main_page_index = 0
extra_hour_row = 40
extra_hour_column = 9
vacation_days_row = 40
vacation_days_column = 15


def main(dir):
    """
    write ueberstunden info into into a nice markdown file

    :param dir: directory where the excels lie and the markdown file will land
    """

    # extract months, hours and vacation days from the excels
    hour_list = {}
    for file_name in os.listdir(dir):
        if "xlsm" in file_name and "$" not in file_name:
            wb = xlrd.open_workbook(dir + "/" + file_name)

            # the important data is on the first sheet
            sheet = wb.sheet_by_index(main_page_index)
            hours = sheet.cell(extra_hour_row, extra_hour_column)
            vacation_days = sheet.cell(vacation_days_row, vacation_days_column)

            # the returned date tuple is a list of integers, build like [year, month, day]
            date_tuple = xlrd.xldate_as_tuple(sheet.cell(2, 10).value, wb.datemode)
            current_year = str(date_tuple[0])

            # to prevent initializing the dict before hand, we check if an entry for the current year exists and if not
            # we add one
            if current_year not in hour_list.keys():
                hour_list[current_year] = []
            hour_list[current_year].append((
                months[date_tuple[1] - 1] + " " + current_year,
                hours.value,
                vacation_days.value
            ))

    # write the data into the mardwon file
    output_file = dir + "/" + "Stundendifferenz.md"
    with open(output_file, "w") as md_file:
        for current_year in hour_list:
            hour_sum = 0
            vacation_sum = 0

            lines = ["# Interessante Daten für " + current_year + "\n",
                     "\n",
                     "| Monat | Stunden | Urlaubstage |\n",
                     "| --- | --- | --- |\n"]

            for month, hours, vacation_days in hour_list[current_year]:
                lines.append("| " + month + " | " + str(hours) + " | " + str(vacation_days) + " |\n")
                hour_sum += hours
                vacation_sum += vacation_days

            lines.append("\n")
            lines.append("**Gesamt Stunden-> " + str(hour_sum) + "** \n")
            lines.append("\n")
            lines.append("**Gesamt Urlaubstage-> " + str(vacation_sum) + "** \n")
            lines.append("\n")
            lines.append("\n")

            md_file.writelines(lines)

    # print some info to please the user
    print("Wrote into", output_file)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("too few arguments, typical usage: ")
        print("    python3 hour_counter.py <zettel_directory>")
        exit(1)
    main(sys.argv[1])
