import csv

with open('baseline.csv', newline='') as f:
    reader = csv.reader(f)
    baseline = list(reader)

baseline = baseline[1:]
# print(baseline[0], baseline[1], baseline[2])

with open('out_ernie.csv', newline='') as f:
    reader = csv.reader(f)
    result = list(reader)

result = result[1:]

old_result = result
result = []
for row in old_result:
    if row[2] != '':
        continue
    result.append(row)

# print(result[0], result[1], result[2])

START = 11
END = 12

min_len = min(len(result), len(baseline))

good = 0
bad = 0

for i in range(min_len):
    row_b = baseline[i]
    row_r = result[i]
    if ((row_b[START] == row_r[START] or row_b[START] == '') and \
            (row_b[END] == row_r[END] or row_b[END] == '')) or \
            (row_b[START] in row_r[START] and row_b[END] in row_r[END]):
        good += 1
    else:
        bad += 1
        print(f"\n#{i}")
        print("baseline:\t", row_b[8])
        print("result:  \t", row_r[8])
        print("baseline:\t", row_b[START], row_b[END])
        print("result:  \t", row_r[START], row_r[END])

print()
print("good:\t", good)
print("bad: \t", bad)
