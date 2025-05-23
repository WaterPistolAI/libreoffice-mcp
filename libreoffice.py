import uno
from com.sun.star.connection import NoConnectException
from com.sun.star.uno import RuntimeException
from com.sun.star.sheet import XSpreadsheetDocument, XDataPilotTables, XDataPilotDescriptor
from com.sun.star.text import XTextDocument
from com.sun.star.table import CellContentType
from com.sun.star.chart import XChartDocument
from com.sun.star.beans import PropertyValue
from com.sun.star.util import SortField, SortDescriptor
from com.sun.star.sheet import TableFilterField
from com.sun.star.sdb import XOfficeDatabaseDocument
from com.sun.star.sdbc import XDataSource, XConnection, XStatement, XResultSet
from com.sun.star.sheet import XConditionalFormat, XConditionEntry
from com.sun.star.drawing import XShape
from mcp.server.fastmcp import FastMCP, Context
from contextlib import asynccontextmanager

UnoRuntime = uno.getComponentContext().ServiceManager.createInstance("com.sun.star.uno.UnoRuntime")

class AppContext:
    def __init__(self, uno_context):
        self.uno_context = uno_context
        self.documents = {}
        self.next_id = 0

    def open_document(self, url):
        desktop = self.uno_context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.uno_context)
        xComp = desktop.loadComponentFromURL(url, "_blank", 0, tuple([]))
        doc_id = str(self.next_id)
        self.next_id += 1
        self.documents[doc_id] = xComp
        return doc_id

    def get_document(self, doc_id):
        return self.documents.get(doc_id)

    def close_document(self, doc_id):
        doc = self.documents.pop(doc_id, None)
        if doc:
            doc.dispose()

def get_uno_context():
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext)
    try:
        context = resolver.resolve("uno:socket,host=localhost,port=2083;urp;StarOffice.ComponentContext")
        return context
    except NoConnectException:
        raise Exception("Error: cannot establish a connection to LibreOffice.")

@asynccontextmanager
async def app_lifespan(server: FastMCP):
    uno_context = get_uno_context()
    app_ctx = AppContext(uno_context)
    try:
        yield app_ctx
    finally:
        for doc in list(app_ctx.documents.values()):
            doc.dispose()

mcp = FastMCP("LibreOffice MCP", lifespan=app_lifespan)

# Existing Tools (assumed from previous implementation)
@mcp.tool()
def open_database(ctx: Context, url: str) -> str:
    app_ctx = ctx.request_context.lifespan_context
    desktop = app_ctx.uno_context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", app_ctx.uno_context)
    xComp = desktop.loadComponentFromURL(url, "_blank", 0, tuple([]))
    doc_id = str(app_ctx.next_id)
    app_ctx.next_id += 1
    app_ctx.documents[doc_id] = xComp
    return doc_id

@mcp.tool()
def create_database(ctx: Context, url: str) -> str:
    app_ctx = ctx.request_context.lifespan_context
    desktop = app_ctx.uno_context.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", app_ctx.uno_context)
    db_doc = desktop.loadComponentFromURL("private:factory/sdatabase", "_blank", 0, tuple([]))
    save_opts = (
        PropertyValue(Name="URL", Value=url),
        PropertyValue(Name="FilterName", Value="StarOffice XML (Base)"),
    )
    db_doc.storeAsURL(url, save_opts)
    doc_id = str(app_ctx.next_id)
    app_ctx.next_id += 1
    app_ctx.documents[doc_id] = db_doc
    return doc_id

@mcp.tool()
def run_query(ctx: Context, doc_id: str, sql: str, username: str = "", password: str = "") -> list[dict] | str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    db_doc = UnoRuntime.queryInterface(XOfficeDatabaseDocument, doc)
    if not db_doc:
        raise RuntimeException("Document is not a database document")
    data_source = db_doc.getDataSource()
    connection = data_source.getConnection(username, password)
    statement = connection.createStatement()
    if sql.lower().strip().startswith("select"):
        result_set = statement.executeQuery(sql)
        meta_data = result_set.getMetaData()
        column_count = meta_data.getColumnCount()
        results = []
        while result_set.next():
            row = {}
            for i in range(1, column_count + 1):
                row[meta_data.getColumnName(i)] = result_set.getString(i)
            results.append(row)
        return results
    else:
        affected_rows = statement.executeUpdate(sql)
        return f"Affected {affected_rows} rows"

@mcp.tool()
def create_table(ctx: Context, doc_id: str, table_name: str) -> str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    db_doc = UnoRuntime.queryInterface(XOfficeDatabaseDocument, doc)
    if not db_doc:
        raise RuntimeException("Document is not a database document")
    data_source = db_doc.getDataSource()
    connection = data_source.getConnection("", "")
    statement = connection.createStatement()
    sql = f"CREATE TABLE {table_name} (ID INTEGER PRIMARY KEY, Name VARCHAR(50))"
    statement.executeUpdate(sql)
    return f"Created table '{table_name}'"

@mcp.tool()
def create_pivot_table(ctx: Context, doc_id: str, sheet_name: str, source_range: str, target_cell: str) -> str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    xSheetDoc = UnoRuntime.queryInterface(XSpreadsheetDocument, doc)
    if not xSheetDoc:
        raise RuntimeException("Document is not a spreadsheet")
    sheets = xSheetDoc.getSheets()
    try:
        sheet = sheets.getByName(sheet_name)
    except Exception:
        raise RuntimeException(f"Sheet '{sheet_name}' not found")
    start_col, start_row = parse_cell_address(source_range.split(':')[0])
    end_col, end_row = parse_cell_address(source_range.split(':')[1])
    source_addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    source_addr.Sheet = sheets.getIndex(sheet_name)
    source_addr.StartColumn = start_col
    source_addr.StartRow = start_row
    source_addr.EndColumn = end_col
    source_addr.EndRow = end_row
    target_col, target_row = parse_cell_address(target_cell)
    target_addr = uno.createUnoStruct("com.sun.star.table.CellAddress")
    target_addr.Sheet = sheets.getIndex(sheet_name)
    target_addr.Column = target_col
    target_addr.Row = target_row
    data_pilot_tables = sheet.getDataPilotTables()
    dp_table = data_pilot_tables.createDataPilotTable(target_addr)
    descriptor = dp_table.getDataPilotDescriptor()
    descriptor.setSourceRange(source_addr)
    row_field = descriptor.getRowFields().createField()
    row_field.setPropertyValue("Orientation", "ROW")
    row_field.setPropertyValue("Field", 0)
    data_field = descriptor.getDataFields().createField()
    data_field.setPropertyValue("Orientation", "DATA")
    data_field.setPropertyValue("Field", end_col - start_col)
    data_field.setPropertyValue("Function", "SUM")
    dp_table.refresh()
    return f"Created pivot table at {target_cell}"

@mcp.tool()
def sort_range(ctx: Context, doc_id: str, sheet_name: str, range_address: str, sort_column: int, ascending: bool) -> str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    xSheetDoc = UnoRuntime.queryInterface(XSpreadsheetDocument, doc)
    if not xSheetDoc:
        raise RuntimeException("Document is not a spreadsheet")
    sheets = xSheetDoc.getSheets()
    try:
        sheet = sheets.getByName(sheet_name)
    except Exception:
        raise RuntimeException(f"Sheet '{sheet_name}' not found")
    start_col, start_row = parse_cell_address(range_address.split(':')[0])
    end_col, end_row = parse_cell_address(range_address.split(':')[1])
    range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
    sort_desc = range.createSortDescriptor()
    for desc in sort_desc:
        if desc.Name == "SortFields":
            sort_field = SortField()
            sort_field.Field = sort_column
            sort_field.SortAscending = ascending
            desc.Value = [sort_field]
    range.sort(sort_desc)
    return f"Sorted range {range_address} by column {sort_column} {'ascending' if ascending else 'descending'}"

@mcp.tool()
def filter_range(ctx: Context, doc_id: str, sheet_name: str, range_address: str, column: int, operator: str, value: str) -> str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    xSheetDoc = UnoRuntime.queryInterface(XSpreadsheetDocument, doc)
    if not xSheetDoc:
        raise RuntimeException("Document is not a spreadsheet")
    sheets = xSheetDoc.getSheets()
    try:
        sheet = sheets.getByName(sheet_name)
    except Exception:
        raise RuntimeException(f"Sheet '{sheet_name}' not found")
    start_col, start_row = parse_cell_address(range_address.split(':')[0])
    end_col, end_row = parse_cell_address(range_address.split(':')[1])
    range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
    filter_desc = range.createFilterDescriptor(True)
    filter_field = TableFilterField()
    filter_field.Field = column
    filter_field.Operator = operator
    filter_field.IsNumeric = value.isdigit()
    filter_field.StringValue = value if not filter_field.IsNumeric else ""
    filter_field.NumericValue = float(value) if filter_field.IsNumeric else 0.0
    filter_desc.setFilterFields([filter_field])
    range.filter(filter_desc)
    return f"Applied filter to range {range_address} on column {column} with {operator} {value}"

@mcp.tool()
def calculate_statistics(ctx: Context, doc_id: str, sheet_name: str, range_address: str) -> dict:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    xSheetDoc = UnoRuntime.queryInterface(XSpreadsheetDocument, doc)
    if not xSheetDoc:
        raise RuntimeException("Document is not a spreadsheet")
    sheets = xSheetDoc.getSheets()
    try:
        sheet = sheets.getByName(sheet_name)
    except Exception:
        raise RuntimeException(f"Sheet '{sheet_name}' not found")
    start_col, start_row = parse_cell_address(range_address.split(':')[0])
    end_col, end_row = parse_cell_address(range_address.split(':')[1])
    range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
    values = []
    for row in range(start_row, end_row + 1):
        for col in range(start_col, end_col + 1):
            cell = sheet.getCellByPosition(col, row)
            if cell.getType() == CellContentType.VALUE:
                values.append(cell.getValue())
    if not values:
        return {"sum": 0, "average": 0}
    total = sum(values)
    average = total / len(values)
    return {"sum": total, "average": average}

# New Base Tools
@mcp.tool()
def list_tables(ctx: Context, doc_id: str) -> list[str]:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    db_doc = UnoRuntime.queryInterface(XOfficeDatabaseDocument, doc)
    if not db_doc:
        raise RuntimeException("Document is not a database document")
    data_source = db_doc.getDataSource()
    connection = data_source.getConnection("", "")
    meta_data = connection.getMetaData()
    result_set = meta_data.getTables(None, None, "%", None)
    tables = []
    while result_set.next():
        tables.append(result_set.getString(3))  # Table name is in column 3
    return tables

@mcp.tool()
def delete_table(ctx: Context, doc_id: str, table_name: str) -> str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    db_doc = UnoRuntime.queryInterface(XOfficeDatabaseDocument, doc)
    if not db_doc:
        raise RuntimeException("Document is not a database document")
    data_source = db_doc.getDataSource()
    connection = data_source.getConnection("", "")
    statement = connection.createStatement()
    sql = f"DROP TABLE {table_name}"
    statement.executeUpdate(sql)
    return f"Deleted table '{table_name}'"

@mcp.tool()
def insert_data(ctx: Context, doc_id: str, table_name: str, data: dict) -> str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    db_doc = UnoRuntime.queryInterface(XOfficeDatabaseDocument, doc)
    if not db_doc:
        raise RuntimeException("Document is not a database document")
    data_source = db_doc.getDataSource()
    connection = data_source.getConnection("", "")
    statement = connection.createStatement()
    columns = ", ".join(data.keys())
    values = ", ".join([f"'{v}'" if isinstance(v, str) else str(v) for v in data.values()])
    sql = f"INSERT INTO {table_name} ({columns}) VALUES ({values})"
    affected_rows = statement.executeUpdate(sql)
    return f"Inserted {affected_rows} row(s) into '{table_name}'"

# New Data Analysis Tools
@mcp.tool()
def create_chart(ctx: Context, doc_id: str, sheet_name: str, data_range: str, chart_type: str, target_cell: str) -> str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    xSheetDoc = UnoRuntime.queryInterface(XSpreadsheetDocument, doc)
    if not xSheetDoc:
        raise RuntimeException("Document is not a spreadsheet")
    sheets = xSheetDoc.getSheets()
    try:
        sheet = sheets.getByName(sheet_name)
    except Exception:
        raise RuntimeException(f"Sheet '{sheet_name}' not found")
    start_col, start_row = parse_cell_address(data_range.split(':')[0])
    end_col, end_row = parse_cell_address(data_range.split(':')[1])
    data_addr = uno.createUnoStruct("com.sun.star.table.CellRangeAddress")
    data_addr.Sheet = sheets.getIndex(sheet_name)
    data_addr.StartColumn = start_col
    data_addr.StartRow = start_row
    data_addr.EndColumn = end_col
    data_addr.EndRow = end_row
    target_col, target_row = parse_cell_address(target_cell)
    chart = sheet.getCharts().addNewByName(f"Chart_{target_cell}", 
                                            uno.createUnoStruct("com.sun.star.awt.Rectangle", 
                                                                target_col * 1000, target_row * 1000, 15000, 10000),
                                            [data_addr])
    chart_doc = chart.getEmbeddedObject()
    chart_doc.getDiagram().setPropertyValue("Type", chart_type.upper())  # e.g., "BAR", "LINE"
    return f"Created {chart_type} chart at {target_cell}"

@mcp.tool()
def apply_conditional_formatting(ctx: Context, doc_id: str, sheet_name: str, range_address: str, condition: str, style: str) -> str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    xSheetDoc = UnoRuntime.queryInterface(XSpreadsheetDocument, doc)
    if not xSheetDoc:
        raise RuntimeException("Document is not a spreadsheet")
    sheets = xSheetDoc.getSheets()
    try:
        sheet = sheets.getByName(sheet_name)
    except Exception:
        raise RuntimeException(f"Sheet '{sheet_name}' not found")
    start_col, start_row = parse_cell_address(range_address.split(':')[0])
    end_col, end_row = parse_cell_address(range_address.split(':')[1])
    range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
    cond_format = range.createConditionalFormat()
    entry = cond_format.createConditionEntryByType("CONDITION")
    entry.setPropertyValue("Formula1", condition)  # e.g., "A1>10"
    entry.setPropertyValue("StyleName", style)
    cond_format.addNew([entry])
    range.setPropertyValue("ConditionalFormat", cond_format)
    return f"Applied conditional formatting to {range_address} with condition '{condition}' and style '{style}'"

@mcp.tool()
def group_range(ctx: Context, doc_id: str, sheet_name: str, range_address: str, by_rows: bool) -> str:
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    xSheetDoc = UnoRuntime.queryInterface(XSpreadsheetDocument, doc)
    if not xSheetDoc:
        raise RuntimeException("Document is not a spreadsheet")
    sheets = xSheetDoc.getSheets()
    try:
        sheet = sheets.getByName(sheet_name)
    except Exception:
        raise RuntimeException(f"Sheet '{sheet_name}' not found")
    start_col, start_row = parse_cell_address(range_address.split(':')[0])
    end_col, end_row = parse_cell_address(range_address.split(':')[1])
    range = sheet.getCellRangeByPosition(start_col, start_row, end_col, end_row)
    if by_rows:
        range.Rows.group(None, True)
    else:
        range.Columns.group(None, True)
    return f"Grouped range {range_address} by {'rows' if by_rows else 'columns'}"

def parse_cell_address(address: str) -> tuple[int, int]:
    if not address or not address[0].isalpha() or not address[1:].isdigit():
        raise RuntimeException(f"Invalid cell address: {address}")
    col = ord(address[0].upper()) - ord('A')
    row = int(address[1:]) - 1
    if col < 0 or row < 0:
        raise RuntimeException(f"Invalid cell address: {address}")
    return col, row