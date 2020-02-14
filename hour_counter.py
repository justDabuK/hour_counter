import xlrd
import os
import sys

months = ["Januar", "Februar", "Maerz", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober",
          "November", "Dezember"]


def main(dir):
    """
    write ueberstunden info into into a nice markdown file

    :param dir: directory where the excels lie and the markdown file will land
    """

    # extract months and hours from the excels
    hour_list = []
    for file_name in os.listdir(dir):
        if "xlsm" in file_name and "$" not in file_name:
            wb = xlrd.open_workbook(dir + "/" + file_name)
            sheet = wb.sheet_by_index(0)
            hours = sheet.cell(40, 9)
            date_tuple = xlrd.xldate_as_tuple(sheet.cell(2, 10).value, wb.datemode)
            hour_list.append((months[date_tuple[1] - 1] + " " + str(date_tuple[0]), hours.value))

    # write the data into the mardwon file
    output_file = dir + "/" + "Stundendifferenz.md"
    hour_sum = 0
    with open(output_file, "w") as md_file:
        lines = ["# Stundendifferenz \n", "\n", "| Monat | Stunden |\n", "| --- | --- |\n"]

        for month, hours in hour_list:
            lines.append("| " + month + " | " + str(hours) + " |\n")
            hour_sum += hours

        lines.append("\n")
        lines.append("**Gesamt -> " + str(hour_sum) + "** \n")

        md_file.writelines(lines)

    # print some info to please the user
    print("Sneak Peak: current status -> ", hour_sum)
    print("Wrote into", output_file)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("too few arguments, typical usage: ")
        print("    python3 hour_counter.py <zettel_directory>")
        exit(1)
    main(sys.argv[1])
