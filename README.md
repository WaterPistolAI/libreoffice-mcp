# **LibreOffice MCP (WIP)**

### Explanation of the Refactor

#### **Why OooDev?**

The OooDev library, as detailed in `api_documentation.md`, provides a Pythonic abstraction over the LibreOffice UNO API, with classes like `CalcDoc`, `WriteDoc`, and utilities for charts, forms, and macros. It simplifies tasks like document creation, cell manipulation, and styling, reducing the need for low-level UNO interface queries (e.g., `UnoRuntime.queryInterface`). OooDev also supports modern Python features (e.g., type hints, context managers) and integrates with extensions like APSO and OooDev.oxt, making it ideal for a 1:1 MCP mapping.

#### **Key Changes from Your Original Implementation**

1. **OooDev Integration**:

   - Replaced direct UNO API calls (e.g., `XSpreadsheetDocument`, `XTextDocument`) with OooDev classes (`CalcDoc`, `WriteDoc`).
   - Used OooDev’s `Lo` class to manage LibreOffice connections, replacing manual `UnoUrlResolver` setup.
   - Leveraged OooDev utilities like `Calc`, `Write`, and `Chart2` for document operations.

2. **Preserved Features**:

   - Kept all your original tools: `open_document`, `close_document`, `get_sheet_names`, `get_cell_value`, `set_cell_value`, `create_new_sheet`, `insert_text`, `create_chart`, and `apply_style`.
   - Enhanced these tools to use OooDev methods (e.g., `CalcDoc.get_sheet_names()` instead of `XSpreadsheetDocument.getSheets()`).
   - Maintained the `parse_cell_address` helper, though OooDev’s `CalcDoc.rng()` method reduces its necessity in some cases.

3. **New Base Tools**:

   - Added `run_query`, `list_tables`, `create_table`, and `insert_data` to support LibreOffice Base operations, using OooDev’s database access methods.
   - These tools align with the `com::sun::star::sdb` namespace from `api.md` but use OooDev’s simplified interfaces.

4. **New Data Analysis Tools**:

   - Added `create_pivot_table`, `sort_range`, and `calculate_statistics` to enhance Calc’s data analysis capabilities.
   - Used OooDev’s `Chart2` module for pivot table creation and `Calc` module for sorting and statistics.

5. **Additional Tools**:
   - Added `new_document` to create new documents of various types, leveraging OooDev’s `create_doc` methods.
   - Added `save_document` to save documents using OooDev’s `save_doc`.
   - Added `run_macro` to execute Python macros, using OooDev’s `MacroLoader` context manager.
   - Added `insert_form_control` to insert form controls in Calc sheets, using OooDev’s `Forms` module.

#### **Alignment with OooDev Features**

- **Document Management**: Uses `WriteDoc`, `CalcDoc`, `DrawDoc` for Writer, Calc, Draw, and Impress documents, covering `com::sun::star::text`, `sheet`, `draw`, and `presentation` namespaces.
- **Base Operations**: Leverages OooDev’s database access for `com::sun::star::sdb` and `sdbc` functionalities.
- **Data Analysis**: Utilizes OooDev’s `Chart2` and `Calc` modules for pivot tables, sorting, and statistics, aligning with `com::sun::star::chart2` and `sheet`.
- **Forms and Macros**: Incorporates OooDev’s `Forms` and `MacroLoader` for `com::sun::star::form` and `script` namespaces.
- **Styling**: Uses OooDev’s style methods (e.g., `Write.style`) for `com::sun::star::style`.

#### **Benefits of the Refactor**

- **Pythonic Interface**: OooDev’s classes and methods are more intuitive and reduce boilerplate compared to UNO’s interface queries.
- **Comprehensive Coverage**: The adapter covers a broad range of OooDev functionalities, from basic document operations to advanced database and analysis tasks.
- **Maintainability**: OooDev’s abstractions simplify maintenance and future expansions (e.g., adding Math or more form controls).
- **Error Handling**: Robust error checking ensures reliability for LLM interactions via MCP.

#### **Limitations and Future Enhancements**

- **Base Document Handling**: The current Base tools use raw UNO interfaces for database operations, as OooDev’s Base support is less developed. Future versions could integrate OooDev’s Base-specific classes if added.
- **Additional Tools**: More OooDev features (e.g., `ooodev.units` for unit conversions, `ooodev.form` for advanced forms, `ooodev.draw` for shapes) could be added as tools.
- **Performance**: OooDev’s abstractions may introduce slight overhead; caching (e.g., `Lo.cache`) could be implemented for frequent operations.

#### **How to Use**

1. **Setup**: Ensure OooDev is installed (`pip install ooo-dev-tools`) and LibreOffice is running with a socket connection (port 2083).
2. **Run the Server**: Start the FastAPI server with this adapter.
3. **Interact via MCP**: Use the MCP client to call tools like `open_document`, `run_query`, or `create_pivot_table`.

- **Coverage Across Applications**: The adapter now supports Writer, Calc, Impress, and Draw, with potential for more Base and Math tools to be added.
- **Rich Feature Set**: It includes not just basic operations (e.g., opening documents, setting cell values) but also advanced features like styling, table insertion, image handling, charting, slide management, and shape drawing.
- **Cross-Cutting Functionality**: Tools like saving, exporting to PDF, and managing document properties apply to all LibreOffice document types, enhancing versatility.
- **Error Handling**: Each tool includes robust checks (e.g., document existence, type validation) to ensure reliability.

## **Old Documetnation**

```
sudo useradd -m -s /bin/bash mcp-libreoffice
sudo mkdir -p /home/mcp-libreoffice/output
sudo chown -R mcp-libreoffice:mcp-libreoffice /home/mcp-libreoffice
```

```
sudo -u mcp-libreoffice soffice --headless --accept="socket,host=localhost,port=${LIBREOFFICE_PORT:-2083};urp;"
```

### **Original Features**:

- **`open_document`**: Opens a document from a URL and assigns it an ID.
- **`close_document`**: Closes a document by its ID.
- **`get_sheet_names`**: Retrieves all sheet names from a Calc spreadsheet.
- **`get_cell_value`**: Gets the value (numeric, text, or formula) of a specified cell in a Calc sheet.
- **`set_cell_value`**: Sets a cell’s value in a Calc sheet, handling both numbers and text.
- **`get_text_content`**: Retrieves the full text content of a Writer document.
- **`insert_text`**: Inserts text at a specific position in a Writer document.
- **`create_chart`**: Creates a chart in a Calc sheet based on a data range and chart type.
- **`parse_cell_address`**: Helper function to convert cell addresses (e.g., "A1") to column and row indices.
- **Spreadsheet Enhancements (Calc)**:
  - **`create_new_sheet`**: Adds a new sheet to a spreadsheet.
  - **`set_cell_formula`**: Sets a formula in a spreadsheet cell (e.g., "=SUM(A1:A10)").
- **Text Document Enhancements (Writer)**:
  - **`insert_table`**: Inserts a table with specified rows and columns at a position.
  - **`apply_style`**: Applies a paragraph style to a text range.
  - **`insert_image`**: Inserts an image from a URL at a specified position.
- **Additional Document Management**:
  - **`save_document`**: Saves a document to a URL with a specified filter (e.g., "writer8" for ODT).
  - **`export_to_pdf`**: Exports a document to PDF format.
  - **`get_document_properties`**: Retrieves metadata like title, author, subject, and keywords.
  - **`set_document_properties`**: Updates document metadata.
- **Presentation Tools (Impress)**:
  - **`insert_slide`**: Inserts a new slide at a specified position in a presentation.
- **Drawing Tools (Draw)**:
  - **`add_shape`**: Adds a shape (e.g., rectangle, circle) to a Draw page with position and size.
- **Macro Execution**:
  - **`run_macro`**: Runs a Basic macro stored in the document.

### Base Tools

The following tools further enhance database management capabilities:

- **`list_tables(doc_id: str) -> list[str]`**: Retrieves a list of all table names in the database.
- **`delete_table(doc_id: str, table_name: str) -> str`**: Deletes a specified table from the database.
- **`insert_data(doc_id: str, table_name: str, data: dict) -> str`**: Inserts a single row of data into a table, with column names and values provided as a dictionary.
- **List Tables**: Use `list_tables("doc_id")` to get all table names in a database.
- **Delete Table**: Call `delete_table("doc_id", "table_name")` to remove a table.
- **Insert Data**: Use `insert_data("doc_id", "table_name", {"ID": 1, "Name": "Test"})` to add a row.

### Data Analysis Tools

These tools extend Calc’s data analysis capabilities:

- **`create_chart(doc_id: str, sheet_name: str, data_range: str, chart_type: str, target_cell: str) -> str`**: Creates a chart (e.g., bar, line) from a data range and places it at the target cell.
- **`apply_conditional_formatting(doc_id: str, sheet_name: str, range_address: str, condition: str, style: str) -> str`**: Applies conditional formatting to a range based on a condition (e.g., "value > 10") and a style.
- **`group_range(doc_id: str, sheet_name: str, range_address: str, by_rows: bool) -> str`**: Groups rows or columns in a range for outlining purposes.
- **Create Chart**: Use `create_chart("doc_id", "Sheet1", "A1:B10", "bar", "C12")` to create a bar chart.
- **Conditional Formatting**: Apply with `apply_conditional_formatting("doc_id", "Sheet1", "A1:A10", "A1>10", "Good")`, assuming "Good" is a defined cell style.

#### **Potential Additions**

- **Base Tools**: Create databases, manage forms, or run queries.
- **Math Tools**: Insert and edit mathematical formulas.
- **Formatting Tools**: Adjust font size, color, or alignment in Writer or Impress.
- **Animation/Transition Tools**: Add animations to Impress slides.
- **Data Analysis Tools**: Implement pivot tables or sorting in Calc.
