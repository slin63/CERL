# Given some .csv, checks rows under a header. If row's value is zero, entire row is deleted.

# 3 args:
#   File name
#   Line with header
#   Column to check for zeroes
import csv
import argparse

from io import StringIO
from itertools import islice

def drop_zeroes(file_name, header_l, col_name):
    with open(file_name) as csvfile:
        # Skips the none-header lines
        for line in xrange(header_l - 1):
            next(csvfile)

        out_name = remove_extensions(file_name) + 'PROCESSED.csv'
        output = open(out_name, 'wb')
        dict_reader = csv.DictReader(csvfile) # Dict reader with proper header loaded (hopefully)
        dict_writer = csv.DictWriter(output, fieldnames=dict_reader.fieldnames)
        stats = _drop_zeroes(dict_reader, dict_writer, col_name, insert_space=True)

        report_stats(stats, out_name, col_name)


def report_stats(stats, out_name, col_name):
    print "File output to: {}\n\tPreserved {} lines out of {} total parsed lines from column: \"{}\"    .".format(out_name, stats[0], stats[1], col_name)


def _drop_zeroes(dict_reader, dict_writer, col_name, insert_space=True):
    # Actually does the work of returning a .csv with no zero rows.
    # http://stackoverflow.com/questions/29971718/reading-both-raw-lines-and-dicionaries-from-csv-in-python
    cleaned_csv = []
    dict_writer.writeheader()
    non_empty_count = 0
    num_lines = 0

    for row in dict_reader:
        num_lines += 1
        if float(row[col_name]) != 0:
            non_empty_count += 1
            dict_writer.writerow(row)

    return non_empty_count, num_lines


def remove_extensions(s):
    period = s.find('.')
    return s[:period]


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Zero row dropper.')
    parser.add_argument('-file', type=str, help="Name of the CSV file.")
    parser.add_argument('-header', type=int, help="Line which contains the header.")
    parser.add_argument('-col', type=str, help="Column to check for zeroes.")

    args = parser.parse_args()

    out_csv = drop_zeroes(args.file, args.header, args.col)

