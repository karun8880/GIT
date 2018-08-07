import xlwt
import xlrd
import os
import datetime 

def empty_row(cs ,row):
	num_cols= cs.ncols
	count_empty=0
	for col_index in range(0,num_cols):
            # get cell value
            cell_val= cs.cell(row, col_index).value
            # check if cell is empty
            if cell_val== '': 
                # set empty cell is True
                empty_cell = True
                # increment counter
                count_empty+= 1
            else:
                # set empty cell is false
                empty_cell= False

            # check if cell is not empty
            #if not empty_cell:
                # print value of cell
                # print('Col #: {} | Value: {}'.format(col_index, cell_val))

        # check the counter if is = num_cols means the whole row is empty       
        if count_empty == num_cols:
		return True
	else:
		return False 

wkbk = xlwt.Workbook()

ch = raw_input( "If header is there press Y or n\t")
sheet_name = raw_input("Enter the sheet name \t")
format2 = xlwt.XFStyle()
format2.num_format_str = 'dd/mm/yyyy'
xlsfiles =  [each for each in  os.listdir(".") if each.endswith("xlsx")]
print "Total files are ", xlsfiles
outsheet = wkbk.add_sheet(sheet_name)

outrow_idx = 0
for f in xlsfiles:
	print "Processing ",f
	try:
		start_value = 0
		if ch.lower() == 'y':
			start_value = 1
		try:
			insheet = xlrd.open_workbook(f).sheet_by_name(sheet_name)
		except:
			insheet=xlrd.open_workbook(f).sheet_by_name(sheet_name+' ')
		print('('+str(insheet.nrows)+","+str(insheet.ncols)+')')
		for row_idx in xrange(start_value,insheet.nrows):
			if empty_row(insheet,row_idx):
				continue
			for col_idx in xrange(insheet.ncols):
				if insheet.cell(row_idx,col_idx).ctype!=3:
					outsheet.write(outrow_idx, col_idx,insheet.cell_value(row_idx, col_idx))
				else:
					outsheet.write(outrow_idx, col_idx,insheet.cell_value(row_idx, col_idx),format2)
			outrow_idx += 1
	except Exception as e :
		print e 
	else :
		print "finished ",f
filename=sheet_name+".xls"
wkbk.save(filename)

