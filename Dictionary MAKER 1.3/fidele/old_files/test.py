import sys
from docx import Document # may need to install this package! 
import xlrd
import xlwt 
import xlsxwriter

#import program1
import program2

input_file = sys.argv[1]
freq_file = sys.argv[2]

# UNIT TESTS
# Program 1
'''
freq_table = xlrd.open_workbook('frequencyTable.xlsx') 
freq_sheet = freq_table.sheet_by_index(0)

freqCount = program1.freqCount
'''

# Check that redundant lines are underlined
# Eyeball test: Choose a random page and ensure that 
#   - the words on the original doc and the generated doc match up
#   - redundant lines are underlined

# Check frequency numbers are correct
# 'abandon' (4), 'Bless you!' (2), 'break down' (3), 'breathe one’s last' (1), 'breath-taking' (1), 'coloring leaves or barks used to strengthen fishing lines' (1), 'East' (1), 'aire' (0)
'''def test_freq(word, freq):
    if freqCount[word] != freq:
        print('freqCount of ', word, ' is ', freqCount[word], ' but should be ', freq)
    return

print('Checking freqency numbers are correct...')
test_freq('abandon', 4)
test_freq('Bless you!', 2)
test_freq('break down', 3)
test_freq('breathe one’s last', 1)
test_freq('breath-taking', 1)
test_freq('coloring leaves or barks used to strengthen fishing lines', 1)
test_freq('East', 1)
if 'aire' in freqCount:
    print('The freqCount of aire is ', freqCount['aire'], ' but aire should not be in freqCount')

# Check ordering of word list is correct
# 'start' (9), 'supply' (4), 'scrutinize' (2), 'superstitous person' (1)
# choose a few series of numbers 
print('Checking ordering of word list is correct...')
def test_ordering(num_list):
    old_i = 0
    old_freq = freq_sheet.cell_value(old_i,1)
    for i in num_list:
        if freq_sheet.cell_value(i, 1) > old_freq:
            print('The ', i, 'th cell (word ', freq_sheet.cell_value(i, 0), ') has cell value ', freq_sheet.cell_value(i, 1), ' which is larger than the ', old_i, 'th cell (word ',  freq_sheet.cell_value(old_i, 0), ') which has cell value ', freq_sheet.cell_value(old_i, 1))
        else:
            old_freq = freq_sheet.cell_value(i,1)
            old_i = i
    return
test_ordering([3914, 6405, 6069, 11301, 20025, 20494])
test_ordering([1,2,3,4,5,6])
'''

# Program 2
ref_table = xlrd.open_workbook('referenceTable.xlsx', formatting_info=True)
ref_sheet = freq_table.sheet_by_index(0)

''' For each, we check lines with one unit and with multiple units
'''

# Check that each line has only one underlined unit

# Check that the unit in column B is the most frequent unit

# Check that the first unit is underlined

# Check that the rest of the units are in alphabetical order

# Check that the reference unit is repeated in column B
