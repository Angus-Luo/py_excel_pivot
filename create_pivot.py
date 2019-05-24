import win32com.client as win32
from win32com.client import DispatchEx
from pandas import read_excel

class create_pivot:
    '''
    Create pivot table from existed Excel file
    '''
    def __init__(self, xlfile, sourcedata_sh_name, pt_sh_name, pt_name):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.win32c = win32.constants
        self.xlfile = xlfile
        self.sourcedata_sh_name = sourcedata_sh_name
        self.pt_sh_name = pt_sh_name
        self.pt_name = pt_name

    def openWorkbook(self):
        try:        
            xlwb = self.excel.Workbooks(self.xlfile)            
        except Exception as e:
            try:
                xlwb = self.excel.Workbooks.Open(self.xlfile)
            except Exception as e:
                print(e)
                xlwb = None                    
        return(xlwb)
    
    def read_data(self):
        data = read_excel(self.xlfile, header=None,sheet_name=self.sourcedata_sh_name)
        data_values = data.values
        return data_values
    
    def set_pivot(self, PageField,RowField, ColumnField, DataField):
        wb = self.openWorkbook()
        data_values = self.read_data()
        Sheet1 = wb.Worksheets(self.sourcedata_sh_name)
        cl1 = Sheet1.Cells(1,1)
        cl2 = Sheet1.Cells(1+len(data_values)-1,1+len(data_values[0])-1)
        PivotSourceRange = Sheet1.Range(cl1,cl2)
        
        wb.Sheets.Add(After=wb.Sheets(1))
        wb.Worksheets[2].Name = self.pt_sh_name
        Sheet2 = wb.Worksheets(2)
        cl3=Sheet2.Cells(4,1)
        PivotTargetRange=  Sheet2.Range(cl3,cl3)
        PivotTableName = self.pt_name

        PivotCache = wb.PivotCaches().Create(SourceType=self.win32c.xlDatabase, SourceData=PivotSourceRange, Version=self.win32c.xlPivotTableVersion14)

        PivotTable = PivotCache.CreatePivotTable(TableDestination=PivotTargetRange, TableName=PivotTableName, DefaultVersion=self.win32c.xlPivotTableVersion14)

        for i in PageField:
            PivotTable.PivotFields(i).Orientation = self.win32c.xlPageField
            PivotTable.PivotFields(i).Position = PageField.index(i) + 1
            PivotTable.PivotFields(i).CurrentPage = 'All'
        
        for i in RowField:
            PivotTable.PivotFields(i).Orientation = self.win32c.xlRowField
            PivotTable.PivotFields(i).Position = 1
        
        for i in ColumnField:
            PivotTable.PivotFields(i).Orientation = self.win32c.xlColumnField
            PivotTable.PivotFields(i).Position = 1
            PivotTable.PivotFields(i).Subtotals = [False for i in range(12)]
        
        for i in DataField:
            DataField = PivotTable.AddDataField(PivotTable.PivotFields(i))
            DataField.NumberFormat = '[BLUE]#,##0;[RED]#,##0'
            
        Sheet2.Shapes.AddChart2(201,4,1,200)
        Sheet2.Shapes.AddChart2(201,3,400,200)
        self.excel.Visible = 1
        self.excel.ActiveWindow.DisplayGridlines = False
        wb.Save()
        self.excel.Application.Quit()
