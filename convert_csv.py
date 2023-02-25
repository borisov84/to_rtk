import csv
import time


start = time.time()


def convert():
    filename = 'rtk.csv'
    result = []
    with open(filename, 'r', encoding='utf-8') as csv_file:
        reader = csv.reader(csv_file)
        for row in reader:
            if str(row).find('Итого т.') != -1:
                a = row[0].split(';')
                if len(row) > 1:
                    # print(row[])
                    a.append(row[1].split(';')[0])
                if a[11] == '':
                    a[11] = '0'
                result.append([a[0].replace('Итого т.', ''), int(a[10]) + int(a[11])/100])
    print(len(result))
    csv_file.close()
    end = time.time() - start
    print(end)


if __name__ == '__main__':
    convert()
