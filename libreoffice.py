from typing import List, Dict
from ooodev.loader import Lo
from ooodev.loader.inst.options import Options
from ooodev.calc import CalcDoc
from ooodev.office.write import Write
from ooodev.write import WriteDoc
from ooodev.draw import DrawDoc
from ooodev.office.chart2 import Chart2
from ooodev.utils.kind.chart2_types import ChartTypes
from ooodev.utils.kind.zoom_kind import ZoomKind
from ooodev.units import UnitMM
from ooodev.form.forms import Forms
from mcp.server.fastmcp import FastMCP, Context
from contextlib import asynccontextmanager
import logging
from dotenv import load_dotenv
import os

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Load environment variables from .env
load_dotenv()

class AppContext:
    """Manages OooDev loader and document state."""
    def __init__(self):
        self.loader = None
        self.documents = {}  # Maps doc_id to document objects (e.g., CalcDoc, WriteDoc)
        self.next_id = 0
        self.output_dir = os.getenv("LIBREOFFICE_OUTPUT_DIR", "/home/mcp-libreoffice/output")
        os.makedirs(self.output_dir, exist_ok=True)

    def start_office(self):
        """Initialize LibreOffice connection via OooDev."""
        if self.loader is None:
            try:
                self.loader = Lo.load_office(
                    connector=Lo.ConnectSocket(host="localhost", port=int(os.getenv("LIBREOFFICE_PORT", "2083"))),
                    opt=Options(log_level="INFO")
                )
            except Exception as e:
                logger.error(f"Failed to connect to LibreOffice: {e}")
                raise
        return self.loader

    def get_document(self, doc_id: str):
        """Retrieve a document by ID."""
        return self.documents.get(doc_id)

    def add_document(self, doc_id: str, doc):
        """Store a document with a unique ID."""
        self.documents[doc_id] = doc

    def remove_document(self, doc_id: str):
        """Remove a document from the store."""
        self.documents.pop(doc_id, None)

    def close_office(self):
        """Close the LibreOffice instance."""
        if self.loader is not None:
            Lo.close_office()
            self.loader = None

    def format_table(self, doc_id: str, sheet_name: str, range_address: str, border_width: int, background_color: str):
        """Format a table range in a Calc document with borders and background color."""
        doc = self.get_document(doc_id)
        if not doc or not isinstance(doc, CalcDoc):
            raise RuntimeException("Document is not a spreadsheet")
        sheet = doc.sheets.get_by_name(sheet_name)
        rng = sheet.rng(range_address)
        # Apply borders
        border = rng.get_border()
        border.set_all_width(UnitMM(border_width))
        rng.set_border(border)
        # Apply background color (e.g., "#FFFF00" for yellow)
        rng.set_background_color(background_color)
        return f"Formatted table {range_address} with border width {border_width} and background {background_color}"

    def create_chart(self, doc_id: str, sheet_name: str, range_address: str, target_cell: str, chart_type: str):
        """Create a chart in a Calc document."""
        doc = self.get_document(doc_id)
        if not doc or not isinstance(doc, CalcDoc):
            raise RuntimeException("Document is not a spreadsheet")
        sheet = doc.sheets.get_by_name(sheet_name)
        chart_types = {
            "column": ChartTypes.Column.TEMPLATE_STACKED.COLUMN,
            "bar": ChartTypes.Bar.TEMPLATE_STACKED.BAR,
            "line": ChartTypes.Line.TEMPLATE_LINE.LINE,
            "pie": ChartTypes.Pie.TEMPLATE_DONUT.PIE
        }
        if chart_type not in chart_types:
            raise RuntimeException(f"Invalid chart type. Use: {', '.join(chart_types.keys())}")
        chart = sheet.charts.insert_chart(
            rng_obj=sheet.rng(range_address),
            cell_name=target_cell,
            width=15,
            height=11,
            diagram_name=chart_types[chart_type]
        )
        return f"Created {chart_type} chart at {target_cell}"

@asynccontextmanager
async def app_lifespan(server: FastMCP):
    app_ctx = AppContext()
    try:
        app_ctx.start_office()
        yield app_ctx
    except Exception as e:
        logger.error(f"Error in LibreOffice lifespan: {e}")
        raise
    finally:
        for doc_id in list(app_ctx.documents.keys()):
            doc = app_ctx.get_document(doc_id)
            if doc:
                doc.close_doc()
            app_ctx.remove_document(doc_id)
        app_ctx.close_office()

# Plugin-specific MCP server
mcp = FastMCP("LibreOffice OooDev MCP", lifespan=app_lifespan)

# Core Document Management Tools
@mcp.tool()
def open_document(ctx: Context, url: str, doc_type: str) -> str:
    """Open an existing LibreOffice document from a URL.

    Args:
        url (str): File path to the document (e.g., '/path/to/doc.odt').
        doc_type (str): Type of document ('writer', 'calc', 'draw', 'impress', 'base').

    Returns:
        str: Document ID for future reference.
    """
    app_ctx = ctx.request_context.lifespan_context
    doc_types = {
        "writer": WriteDoc,
        "calc": CalcDoc,
        "draw": DrawDoc,
        "impress": DrawDoc,
        "base": None
    }
    if doc_type not in doc_types:
        raise RuntimeException(f"Invalid document type. Use: {', '.join(doc_types.keys())}")
    try:
        if doc_type == "base":
            doc = Lo.open_doc(fnm=url, loader=app_ctx.loader)
        else:
            doc_class = doc_types[doc_type]
            doc = doc_class.from_path(fnm=url, lo_inst=app_ctx.loader)
        doc_id = f"doc_{app_ctx.next_id}"
        app_ctx.next_id += 1
        app_ctx.add_document(doc_id, doc)
        return doc_id
    except Exception as e:
        raise RuntimeException(f"Failed to open document: {str(e)}")

@mcp.tool()
def new_document(ctx: Context, doc_type: str) -> str:
    """Create a new LibreOffice document of the specified type.

    Args:
        doc_type (str): Type of document ('writer', 'calc', 'draw', 'impress', 'base').

    Returns:
        str: Document ID for future reference.
    """
    app_ctx = ctx.request_context.lifespan_context
    doc_types = {
        "writer": WriteDoc,
        "calc": CalcDoc,
        "draw": DrawDoc,
        "impress": DrawDoc,
        "base": None
    }
    if doc_type not in doc_types:
        raise RuntimeException(f"Invalid document type. Use: {', '.join(doc_types.keys())}")
    try:
        if doc_type == "base":
            doc = Lo.create_doc(doc_type="sbase", loader=app_ctx.loader)
        else:
            doc_class = doc_types[doc_type]
            doc = doc_class.create_doc(lo_inst=app_ctx.loader)
        doc_id = f"doc_{app_ctx.next_id}"
        app_ctx.next_id += 1
        app_ctx.add_document(doc_id, doc)
        return doc_id
    except Exception as e:
        raise RuntimeException(f"Failed to create new document: {str(e)}")

@mcp.tool()
def save_document(ctx: Context, doc_id: str, url: str) -> str:
    """Save a document to a specified URL.

    Args:
        doc_id (str): Document ID.
        url (str): File path to save the document.

    Returns:
        str: Confirmation message.
    """
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    try:
        doc.save_doc(fnm=os.path.join(app_ctx.output_dir, url))
        return f"Document saved to {url}"
    except Exception as e:
        raise RuntimeException(f"Failed to save document: {str(e)}")

@mcp.tool()
def close_document(ctx: Context, doc_id: str) -> str:
    """Close a LibreOffice document.

    Args:
        doc_id (str): Document ID.

    Returns:
        str: Confirmation message.
    """
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    try:
        doc.close_doc()
        app_ctx.remove_document(doc_id)
        return f"Document {doc_id} closed"
    except Exception as e:
        raise RuntimeException(f"Failed to close document: {str(e)}")

# Calc (Spreadsheet) Tools
@mcp.tool()
def get_sheet_names(ctx: Context, doc_id: str) -> List[str]:
    """Get a list of sheet names in a spreadsheet document."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, CalcDoc):
        raise RuntimeException("Document is not a spreadsheet")
    return doc.get_sheet_names()

@mcp.tool()
def get_cell_value(ctx: Context, doc_id: str, sheet_name: str, cell_address: str) -> str:
    """Get the value of a cell in a spreadsheet."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, CalcDoc):
        raise RuntimeException("Document is not a spreadsheet")
    sheet = doc.sheets.get_by_name(sheet_name)
    cell = sheet[cell_address]
    if cell.is_empty():
        return ""
    return str(cell.value)

@mcp.tool()
def set_cell_value(ctx: Context, doc_id: str, sheet_name: str, cell_address: str, value: str) -> str:
    """Set the value of a cell in a spreadsheet."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, CalcDoc):
        raise RuntimeException("Document is not a spreadsheet")
    sheet = doc.sheets.get_by_name(sheet_name)
    try:
        cell = sheet[cell_address]
        cell.value = float(value)
    except ValueError:
        cell.value = value
    return f"Set {cell_address} to {value}"

@mcp.tool()
def create_new_sheet(ctx: Context, doc_id: str, sheet_name: str) -> str:
    """Create a new sheet in a spreadsheet document."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, CalcDoc):
        raise RuntimeException("Document is not a spreadsheet")
    doc.sheets.insert_new_by_name(sheet_name, len(doc.get_sheet_names()))
    return f"Created new sheet '{sheet_name}'"

# Data Analysis Tools (Calc)
@mcp.tool()
def create_pivot_table(ctx: Context, doc_id: str, sheet_name: str, source_range: str, target_cell: str) -> str:
    """Create a pivot table from a data range in a spreadsheet."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, CalcDoc):
        raise RuntimeException("Document is not a spreadsheet")
    sheet = doc.sheets.get_by_name(sheet_name)
    tbl_chart = sheet.charts.insert_chart(
        rng_obj=sheet.rng(source_range),
        cell_name=target_cell,
        width=15,
        height=11,
        diagram_name=ChartTypes.Pivot.TEMPLATE_PIVOT.PIVOT
    )
    return f"Created pivot table at {target_cell}"

@mcp.tool()
def sort_range(ctx: Context, doc_id: str, sheet_name: str, range_address: str, sort_column: int, ascending: bool) -> str:
    """Sort a range in a spreadsheet by a specified column."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, CalcDoc):
        raise RuntimeException("Document is not a spreadsheet")
    sheet = doc.sheets.get_by_name(sheet_name)
    rng = sheet.rng(range_address)
    sort_field = SortField()
    sort_field.Field = sort_column
    sort_field.SortAscending = ascending
    rng.sort([sort_field])
    return f"Sorted range {range_address} by column {sort_column} {'ascending' if ascending else 'descending'}"

@mcp.tool()
def calculate_statistics(ctx: Context, doc_id: str, sheet_name: str, range_address: str) -> Dict[str, float]:
    """Calculate basic statistics (sum, average) for a range in a spreadsheet."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, CalcDoc):
        raise RuntimeException("Document is not a spreadsheet")
    sheet = doc.sheets.get_by_name(sheet_name)
    rng = sheet.rng(range_address)
    values = [cell.value for cell in rng if isinstance(cell.value, (int, float))]
    if not values:
        return {"sum": 0.0, "average": 0.0}
    total = sum(values)
    average = total / len(values)
    return {"sum": total, "average": average}

# Base (Database) Tools
@mcp.tool()
def run_query(ctx: Context, doc_id: str, sql: str, username: str = "", password: str = "") -> List[Dict[str, str]] | str:
    """Execute an SQL query on a Base database."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    data_source = doc.getDataSource()
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
def list_tables(ctx: Context, doc_id: str) -> List[str]:
    """List all tables in a Base database."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    data_source = doc.getDataSource()
    connection = data_source.getConnection("", "")
    meta_data = connection.getMetaData()
    result_set = meta_data.getTables(None, None, "%", None)
    tables = []
    while result_set.next():
        tables.append(result_set.getString(3))
    return tables

@mcp.tool()
def create_table(ctx: Context, doc_id: str, table_name: str, columns: List[Dict[str, str]]) -> str:
    """Create a new table in a Base database."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    data_source = doc.getDataSource()
    connection = data_source.getConnection("", "")
    statement = connection.createStatement()
    column_defs = ", ".join(f"{col['name']} {col['type']}" for col in columns)
    sql = f"CREATE TABLE {table_name} ({column_defs})"
    statement.executeUpdate(sql)
    return f"Created table '{table_name}'"

@mcp.tool()
def insert_data(ctx: Context, doc_id: str, table_name: str, data: Dict[str, str]) -> str:
    """Insert a row of data into a Base table."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    data_source = doc.getDataSource()
    connection = data_source.getConnection("", "")
    statement = connection.createStatement()
    columns = ", ".join(data.keys())
    values = ", ".join([f"'{v}'" if isinstance(v, str) else str(v) for v in data.values()])
    sql = f"INSERT INTO {table_name} ({columns}) VALUES ({values})"
    affected_rows = statement.executeUpdate(sql)
    return f"Inserted {affected_rows} row(s) into '{table_name}'"

@mcp.tool()
def create_form(ctx: Context, doc_id: str, table_name: str, form_name: str) -> str:
    """Create a form in a Base database linked to a table."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    forms = Forms(doc=doc)
    form = forms.insert_form(name=form_name)
    form.setPropertyValue("ContentType", "Table")
    form.setPropertyValue("Command", table_name)
    return f"Created form '{form_name}' linked to table '{table_name}'"

@mcp.tool()
def create_report(ctx: Context, doc_id: str, table_name: str, report_name: str) -> str:
    """Create a report in a Base database based on a table."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    report_designer = Lo.create_instance_mcf("com.sun.star.report.pentaho.SOReportJobFactory", loader=app_ctx.loader)
    report = report_designer.createReport()
    report.setPropertyValue("Command", table_name)
    report.setPropertyValue("Caption", report_name)
    return f"Created report '{report_name}' based on table '{table_name}'"

# Writer Tools
@mcp.tool()
def insert_text(ctx: Context, doc_id: str, text: str, position: int) -> str:
    """Insert text into a Writer document at a specified position."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, WriteDoc):
        raise RuntimeException("Document is not a text document")
    cursor = doc.get_cursor()
    total_length = len(Write.get_text_string(cursor))
    if position < 0 or position > total_length:
        raise RuntimeException(f"Position {position} out of range (0-{total_length})")
    cursor.goto_start(False)
    cursor.go_right(position, False)
    Write.append(cursor, text)
    return f"Inserted '{text}' at position {position}"

@mcp.tool()
def apply_style(ctx: Context, doc_id: str, style_name: str, start: int, end: int) -> str:
    """Apply a paragraph style to a text range in a Writer document."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, WriteDoc):
        raise RuntimeException("Document is not a text document")
    cursor = doc.get_cursor()
    cursor.goto_start(False)
    cursor.go_right(start, False)
    cursor.go_right(end - start, True)
    Write.style(cursor, style_name)
    return f"Applied style '{style_name}' to text from position {start} to {end}"

# Additional Tools
@mcp.tool()
def run_macro(ctx: Context, doc_id: str, macro_name: str) -> str:
    """Run a Python macro in a document."""
    from ooodev.macro.macro_loader import MacroLoader
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc:
        raise RuntimeException("Document not found")
    with MacroLoader():
        script = doc.getScriptProvider().getScript(f"vnd.sun.star.script:{macro_name}?language=Python&location=document")
        script.invoke((), (), ())
    return f"Executed macro '{macro_name}'"

@mcp.tool()
def insert_form_control(ctx: Context, doc_id: str, sheet_name: str, cell_address: str, control_type: str, label: str) -> str:
    """Insert a form control into a Calc sheet."""
    app_ctx = ctx.request_context.lifespan_context
    doc = app_ctx.get_document(doc_id)
    if not doc or not isinstance(doc, CalcDoc):
        raise RuntimeException("Document is not a spreadsheet")
    sheet = doc.sheets.get_by_name(sheet_name)
    cell = sheet[cell_address]
    control_types = {
        "checkbox": Forms.insert_control_check_box,
        "button": Forms.insert_control_button,
        "listbox": Forms.insert_control_list_box
    }
    if control_type not in control_types:
        raise RuntimeException(f"Invalid control type. Use: {', '.join(control_types.keys())}")
    control = control_types[control_type](cell=cell, label=label)
    return f"Inserted {control_type} control '{label}' at {cell_address}"

@mcp.tool()
def format_table(ctx: Context, doc_id: str, sheet_name: str, range_address: str, border_width: int, background_color: str) -> str:
    """Format a table range in a Calc document with borders and background color."""
    app_ctx = ctx.request_context.lifespan_context
    return app_ctx.format_table(doc_id, sheet_name, range_address, border_width, background_color)

@mcp.tool()
def create_chart(ctx: Context, doc_id: str, sheet_name: str, range_address: str, target_cell: str, chart_type: str) -> str:
    """Create a chart in a Calc document."""
    app_ctx = ctx.request_context.lifespan_context
    return app_ctx.create_chart(doc_id, sheet_name, range_address, target_cell, chart_type)