import pandas as pd
import math
import os

print(
    '''------------------------------
|  Flag Distance Calculator  |
|         By Archie          |
------------------------------'''
)

# control vars
defualtVals = ''.upper()
while defualtVals not in ['Y', 'N']:
    defualtVals = input("Use default values? (Default: Cunning Running Rich Task) ['y' for Yes, 'n' for No]: ").upper()

if defualtVals == 'Y':
    entryHeight = 40
    exitHeight = 10
    flagCount = 13
    fieldLength = 120
    flagSpacing = 10
else:
    try:
        entryHeight = int(input("Enter the height between the entry and the flag line: "))
        exitHeight = int(input("Enter the height between the exit and the flag line: "))
        flagCount = int(input("Enter the number of flags: "))
        fieldLength = int(input("Enter the length of the field: "))
    except ValueError:
        print("Invalid input.")
        exit()
    flagSpacing = (fieldLength / flagCount)

df = pd.DataFrame(columns=['Flag', 'Length from Entry to Flag', 'Length from Flag to Exit', 'Total Distance'])

for index in range(1, (flagCount + 1)):
    print(str(index) + " ------------")

    t1Length = ((flagSpacing * index) - flagSpacing)
    print(f'T1 Length: {t1Length}')

    t1Hypo = math.sqrt((t1Length ** 2) + (entryHeight ** 2))
    print(f'T1 Hypo: {t1Hypo}')

    t2Length = fieldLength - t1Length
    print(f'T2 Length: {t2Length}')

    t2Hypo = math.sqrt((t2Length ** 2) + (exitHeight ** 2))
    print(f'T2 Hypo: {t2Hypo}')

    outLength = t1Hypo + t2Hypo

    df = pd.concat([df, pd.DataFrame([[index, t1Hypo, t2Hypo, outLength]], columns=['Flag', 'Length from Entry to Flag', 'Length from Flag to Exit', 'Total Distance'])], ignore_index=True)

try:
    df.to_excel('output.xlsx', index=False)
except PermissionError:
    print("Please close the file 'output.xlsx' and try again")

print("Output saved to 'output.xlsx' in the same folder as this exe")
