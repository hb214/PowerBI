import * as agGrid from 'ag-grid-community';
import { GridOptions } from 'ag-grid-community';
import 'ag-grid-community/styles/ag-grid.css';
// import 'ag-grid-community/styles/ag-theme-alpine.css';

import powerbi from 'powerbi-visuals-api';
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
//import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import { valueFormatter } from 'powerbi-visuals-utils-formattingutils';
import VisualObjectInstance = powerbi.VisualObjectInstance;

/**
 * The Matrix Class responsible for creating the grid and formatting via ag-grid.
 */
export class Matrix {
  // The columnDefs array to hold columns
  static columnDefs: any[] = [];

  // The rowData array that holds the data for the rows that are mapped to columnDefs
  static rowData: any[] = [];

  // The powerBI selection API for selection and expansion
  static selectionManager: ISelectionManager;

  // The powerBI host API
  static host: IVisualHost;

  static pinnedTotalRow;

  // The formatting settings
  static formattingSettings;

  // The powerBI dataView provided by the API
  static dataView: any;

  // The rowLevels information provided by the API
  static rowLevels: powerbi.DataViewHierarchyLevel[];

  // The columnLevels information provided byt he API
  static columnLevels: powerbi.DataViewHierarchyLevel[];

  // The row children of the matrix dataview
  static rowChildren: powerbi.DataViewMatrixNode[];

  // The column children of the matrix dataview
  static columnChildren: powerbi.DataViewMatrixNode[];

  // An array where the rowChildrenNodes get added too
  static rowChildrenNodes = [];

  // The array where the previous array gets sorted
  static rowChildrenNodesSorted = [];

  // The array to map nodes for expansion
  static nodesToExpand = [];

  // Boolean value that decides if the array gets sorted, otherwise defaults to expansion upwards
  static expandUpwards = false;

  // A helper array that gets used during the creation of nodes
  static tempRowChildren = [];

  // The persisted properties
  static persistedProperties = undefined;

  /**
   * The grid options!
   */
  static gridOptions: GridOptions = {
    // Grab the columnDefs
    columnDefs: this.columnDefs,
    // Grab the rowData
    rowData: undefined,

    // Default rowHeight is 25
    rowHeight: 25,

    // Default header height is 25
    headerHeight: 25,

    // Default col def properties get applied to all columns
    defaultColDef: {
      resizable: true,
      editable: false,
      suppressMovable: true,
      wrapHeaderText: true,
      autoHeaderHeight: true,
      wrapText: true,
      autoHeight: true,
      minWidth: 0,
    },

    // Prevents the context menu in the browser
    preventDefaultOnContextMenu: true,

    // Default row buffer
    rowBuffer: 10,

    // Animate rows as it looks nicer
    animateRows: false,

    // If this is not enabled, the eventlisteners on the columns will not work when you have a lot of columns
    //that warrants a horizontal scroller. This should not be an issue most of the time if performance suffers
    suppressColumnVirtualisation: true,

    // Does not work with the default theme
    columnHoverHighlight: true,
    suppressRowHoverHighlight: false,

    // When clicking a cell
    onCellClicked: (e) => {
      this.selectOnClick(e);
    },

    // Context menu (Right click)
    onCellContextMenu: (e) => {
      this.selectOnClick(e);
    },

    // When the column is resized, we want to persist the properties
    onColumnResized: (params) => {
      // This source is by dragging and only triggered when the dragging is finished
      if (params.source === 'uiColumnResized' && params.finished) {
        this.persistPropertiesToAPI();
      }
    },

    // When the grid is ready do this
    onGridReady: () => {
      //  Sets the event listeners for the headers
      this.addHeaderEventListeners();

      // Sets the event listeners for the body
      this.addBodyEventListeners();

      // Add the expansion buttons
      this.AddExpandButtons();

      // Format the columns
      this.formatColumns(this.gridOptions);

      // Format all the rows
      this.formatRows();

      // Format the expanded rows
      this.formatExpandedRows();

      // Format rowheaders
      this.formatRowHeaders();

      // Format column headers
      this.formatColHeaders();

      // Format specific columns
      this.formatSpecificColumns();

      // Format specific rows
      this.formatSpecificRows();

      // Format the total row
      this.formatTotal();

      // Format the height change
      this.onFinishHeightChange();

      // Get the persisted properties
      this.getAndSetPersistedProperties();

      // Allow the visibility to be shown by getting the gridDiv and setting the visibility to visible
      const gridDiv = document.getElementById('myGrid');
      gridDiv.style.visibility = 'visible';
    },

    // When the viewport changes (New rows are added/removed)
    onViewportChanged: () => {
      // Add expand buttons and format them
      this.AddExpandButtons();

      // Format all the rows
      this.formatRows();

      // Format the expansion buttons
      this.formatExpandedRows();

      // Format rowheaders
      this.formatRowHeaders();

      // Format specific columns
      this.formatSpecificColumns();

      // Format specific rows
      this.formatSpecificRows();

      // Format the total row
      this.formatTotal();

      // Format the height change
      this.onFinishHeightChange();
    },

    /**
     * END OF GRID OPTIONS
     */
  };

  /**
   * Getting the persisted properties and applying them to the grid
   */
  public static getAndSetPersistedProperties() {
    // Check if the persisted properties are undefined
    if (this.persistedProperties === undefined) {
      return;
    }

    // Get the persisted properties
    const persistedProperties = JSON.parse(
      this.persistedProperties['colWidth']
    );

    // Apply the persisted properties
    this.setPersistedProperties(persistedProperties);
  }

  /**
   * Applying the persisted properties
   */
  public static setPersistedProperties(props) {
    // Check if the persisted properties are undefined
    if (this.persistedProperties === undefined) {
      return;
    }

    // Check if the columnCard enables drag
    if (this.formattingSettings.columnCard.enableDrag.value === false) {
      return;
    }

    // Apply the persisted properties
    for (const prop of props) {
      // Get the column
      const column = this.gridOptions.api.getColumnDef(prop['colId']);

      // If column is undefined, skip
      if (column === null || column === undefined) {
        continue;
      }

      // Ensure that the field matches the persisted properties
      if (column['field'] === prop['field']) {
        // Apply the column width
        this.gridOptions.columnApi.setColumnWidths([
          { key: prop['colId'], newWidth: prop['width'] },
        ]);
      }
    }
  }

  /**
   * The persist properties function to save the column width
   */
  public static persistPropertiesToAPI() {
    // Get all columnDefs
    const allColumnDef = this.gridOptions.api.getColumnDefs();

    // Container
    const widthArray = [];

    // Loop through all of them
    for (const colDef of allColumnDef) {
      // Push object with colId, field and width
      widthArray.push({
        colId: colDef['colId'],
        field: colDef['field'],
        width: colDef['width'],
      });
    }

    // Create the instance
    const instance: VisualObjectInstance = {
      objectName: 'persistedProperties',
      selector: undefined,
      properties: {
        colWidth: JSON.stringify(widthArray),
      },
    };

    // Persist the properties
    this.host.persistProperties({ merge: [instance] });
  }

  /**
   * Populates the matrix via the Power BI API
   */
  public static populateMatrixInformation(
    dataView: powerbi.DataView,
    selectionManager,
    host,
    formattingSettings
  ) {
    // Update selection manager
    this.selectionManager = selectionManager;

    // Update host
    this.host = host;

    // Update dataView
    this.dataView = dataView;

    // Update rowLevels
    this.rowLevels = dataView.matrix.rows.levels;

    // Update columnLevels
    this.columnLevels = dataView.matrix.columns.levels;

    // Update rowChildren
    this.rowChildren = dataView.matrix.rows.root.children;

    // Update columnChildren
    this.columnChildren = dataView.matrix.columns.root.children;

    // Update the formatting settings
    this.formattingSettings = formattingSettings;

    // Update the persisted properties
    if (
      dataView.metadata.objects &&
      dataView.metadata.objects.persistedProperties
    ) {
      this.persistedProperties = dataView.metadata.objects.persistedProperties;
    }

    // Return the formatted matrix
    return this.formatMatrix(dataView);
  }


  private static resetState() {
    // Clear out the child nodes in case they are already populated
    this.rowChildrenNodes = [];
    this.rowChildrenNodesSorted = [];

    // Clear out the nodes to expand
    this.nodesToExpand = [];
  }

  private static extractMetadata(dataView: powerbi.DataView): any {
    // Deconstruct dataview to a matrix
    const matrix = dataView.matrix;

    // Dynamic header on the first column. First we set a let variable due to a try/catch
    let dynamicHeader: string;

    // Searches for the first column with the role of Rows and sets the dynamicHeader to that column's display name
    for (const column of dataView.metadata.columns) {
      if (column.roles['Rows']) {
        dynamicHeader = column.displayName;
      }
    }

    return {
      matrix,
      dynamicHeader,
      rowHeight: this.formattingSettings.rowCard.height.value,
      headerHeight: this.formattingSettings.colHeadersCard.height.value,
    }
  }

  private static applyGridStyling() {
    // Create an object to hold CSS styling from the rowCard formatting settings
    const rowCard = this.formattingSettings.rowCard;
    // Formatting for the rowHeadersCard (The first column)
    const rowHeadersCard = this.formattingSettings.rowHeadersCard;


    const gridCellStyling = {
      justifyContent: rowCard.alignment.value.value,
      color: rowCard.fontColor.value.value,
      fontFamily: rowCard.fontFamily.value,
      fontSize: `${rowCard.fontSize.value}px`,
      fontWeight: rowCard.enableBold.value ? 'bold' : 'normal',
      fontStyle: rowCard.enableItalic.value ? 'italic' : 'normal',
      textIndent: `${rowCard.indentation.value}px`,
    };


    const categoryCellStyling = {
      textIndent: `${rowHeadersCard.indentation.value}px`,
      justifyContent: rowHeadersCard.alignment.value.value,
      fontWeight: rowHeadersCard.enableBold.value ? 'bold' : 'normal',
      fontStyle: rowHeadersCard.enableItalic.value ? 'italic' : 'normal',
      color: rowHeadersCard.fontColor.value.value,
      fontSize: `${rowHeadersCard.fontSize.value}px`,
      fontFamily: rowHeadersCard.fontFamily.value,
      borderRight: rowHeadersCard.enableRightBorder.value
        ? `${rowHeadersCard.borderWidth.value}px ${rowHeadersCard.borderStyle.value.value
        } ${hex2rgba(
          rowHeadersCard.borderColor.value.value,
          rowHeadersCard.borderOpacity.value
        )}`
        : 'none',
      borderTop: rowHeadersCard.enableTopBorder.value
        ? `${rowHeadersCard.borderWidth.value}px ${rowHeadersCard.borderStyle.value.value
        } ${hex2rgba(
          rowHeadersCard.borderColor.value.value,
          rowHeadersCard.borderOpacity.value
        )}`
        : 'none',
      borderBottom: rowHeadersCard.enableBottomBorder.value
        ? `${rowHeadersCard.borderWidth.value}px ${rowHeadersCard.borderStyle.value.value
        } ${hex2rgba(
          rowHeadersCard.borderColor.value.value,
          rowHeadersCard.borderOpacity.value
        )}`
        : 'none',
    };
    return { gridCellStyling, categoryCellStyling, rowHeadersCard };
  }

  private static handleSingularMeasureMatrix(matrix, dataView, columnDefs, colId, dynamicHeader, gridCellStyling, categoryCellStyling, rowHeadersCard) {
    // A singular measure (or multiple measures) without any columns or rows
    if (
      !Object.prototype.hasOwnProperty.call(
        matrix.columns.root,
        'childIdentityFields'
      ) &&
      !Object.prototype.hasOwnProperty.call(
        matrix.rows.root,
        'childIdentityFields'
      )
    ) {
      //console.log('WOWWWW');
      // Loop through the valueSources
      matrix.valueSources.forEach((source) => {
        // Push into columnDefs
        columnDefs.push({
          field: source.displayName,
          colId: colId++,
          cellStyle: gridCellStyling,
          cellClass: 'gridCell',
        });
      });
    }
    else if (
      !Object.prototype.hasOwnProperty.call(matrix.columns.root, 'childIdentityFields') &&
      Object.prototype.hasOwnProperty.call(matrix.rows.root, 'childIdentityFields')
    ) {
      // Insert the first value from the row node
      columnDefs.push({
        field: dynamicHeader,
        colId: colId++,
        cellClass: 'categoryCell',
        cellStyle: rowHeadersCard.enableCard.value
          ? categoryCellStyling
          : gridCellStyling,
      });

      // Loop through the columns in the metadata list. Ensure the loop skips a repeat of the first column.
      dataView.matrix.valueSources.forEach((column) => {
        //  Check if repeated column, if so then return
        if (
          column.expr['ref'] === matrix.rows.root.childIdentityFields[0]['ref']
        ) {
          return;
        }

        // Check if column is on a lower level, if so return as it should not be a column
        if (
          column['roles']['Value'] != true &&
          column['rolesIndex']['Rows'][0] != 0
        ) {
          // Return as not to add it as a column
          return;
        }

        // Otherwise, push the object to the columnDefs
        columnDefs.push({
          field: column.displayName,
          colId: colId++,
          cellStyle: gridCellStyling,
          cellClass: 'gridCell',
        });
      });
    }
  }

  private static defineColumns(matrix, dataView, columnDefs, colId, dynamicHeader, gridCellStyling, categoryCellStyling, rowHeadersCard) {
    // Assuming columns and rows and measures have been inserted (Works without measures and rows)
    /* if (matrix.columns.root.hasOwnProperty('childIdentityFields'))  */
    if (Object.prototype.hasOwnProperty.call(matrix.columns.root, 'childIdentityFields')) {
      // Pushes the leftmost column name
      columnDefs.push({
        field: dynamicHeader,
        colId: colId++,
        cellClass: 'categoryCell',
        cellStyle: rowHeadersCard.enableCard.value
          ? categoryCellStyling
          : gridCellStyling,
      });

      // Iterates through the column children
      matrix.columns.root.children.forEach((column) => {
        columnDefs.push({
          field: String(column.value),
          colId: colId++,
          cellClass: 'gridCell',
          cellStyle: gridCellStyling,
        });
      });
    
      // Pushes a "Total" field as the last field
      const lengthOfColumnDefs = columnDefs.length;

      // This check is neccessary if calculation groups are used. With normal columns it is undefined, otherwise it has a clear label.
      if (
        columnDefs[lengthOfColumnDefs - 1]['field'] === undefined ||
        columnDefs[lengthOfColumnDefs - 1]['field'] === 'Undefined'
      ) {
        // Undefined columns are the Total via the Power BI API
        columnDefs[lengthOfColumnDefs - 1]['field'] = 'Total';

        // Hide the column via the API if toggled off
        if (this.formattingSettings.columnCard.enableTotal.value === false) {
          columnDefs[lengthOfColumnDefs - 1]['hide'] = true;
        }
      }
    }

    // Updates the columnDefs in the object,
    this.columnDefs = columnDefs;
  }

  private static buildRowData(matrix, columnDefs, rowData, dynamicHeader) {
    // Declare the rowObj variable to be pushed into rowData. Resets every loop
    let valueSourcesIndexBackUp = 0;  

    matrix.rows.root.children.forEach((row) => {
      const identityHolder = [];
      const rowObj = {};

      if (row.isSubtotal !== true) {
        identityHolder.push(row);
      }
      rowObj['identity'] = identityHolder;

      if (
        Object.prototype.hasOwnProperty.call(row, 'values') &&
        !Object.prototype.hasOwnProperty.call(matrix.columns.root, 'childIdentityFields') &&
        !Object.prototype.hasOwnProperty.call(matrix.rows.root, 'childIdentityFields')
      ) {
        const rowValues = row.values as {[key: string]: any};
        let index = 0;

        Object.values(rowValues).forEach((value) => {
          const typedValue = value as {value: any; valueSourceIndex?: number};

          const valueFormatted = this.valueFormatterMatrix(
            typedValue.value,
            //value.value,
            typedValue.valueSourceIndex === undefined
              ? valueSourcesIndexBackUp
              : typedValue.valueSourceIndex,
            rowValues
          );

          rowObj[Object(columnDefs)[index]['field']] = valueFormatted;
          valueSourcesIndexBackUp++;
          index++;
        });
      }
      
      // If there are multiple values in a row
      /* else if (row.hasOwnProperty('values') || row.children) */
      else if (Object.prototype.hasOwnProperty.call(row, 'values') || row.children) {
        // Deconstruct row into values
        const rowValues = row.values as {[key: string]: any};

        // Index iterator
        let index = 0;

        // Insert row header
        rowObj[Object(columnDefs)[index]['field']] = row.value;

        // Add the nodes to the nodesToExpand array so expand buttons can be programatically added
        if (
          Object.prototype.hasOwnProperty.call(row, 'isCollapsed') ||
          Object.prototype.hasOwnProperty.call(row, 'identity')
        ) {
          this.nodesToExpand.push(row);
        }

        // Increment index
        index++;

        // Valuesource index back up (For calculation groups)
        let valueSourcesIndexBackUp = 0;
        // Try catch as sometimes it is values and sometimes it is children, but they never exist together in the dataView.
        try {
          // Loop through last level of children and insert into rowObj
          Object.values(rowValues).forEach((value) => {
            const typedValue = value  as {value: any; valueSourceIndex?: number};
            // Come back
            const valueFormatted = this.valueFormatterMatrix(
              typedValue.value,
              typedValue.valueSourceIndex === undefined
                ? valueSourcesIndexBackUp
                : typedValue.valueSourceIndex,
              rowValues
            );

            rowObj[Object(columnDefs)[index]['field']] = valueFormatted;

            // Increment the index counter
            index++;
            valueSourcesIndexBackUp++;
          });

          // If catch then it is children and we need to recursively loop through them
        } catch {
          // If expand upwards is enabled, then pop it and insert it after loop
          let lastItem = null;

          if (this.formattingSettings.expansionCard.expandUp.value === true) {
            lastItem = this.nodesToExpand.pop();
          }

          // Loop through children and insert into rowObj
          row.children.forEach((child) => {
            this.tempRowChildren.length = 0;
            this.traverseChildNodes(
              child,
              columnDefs,
              row.value,
              identityHolder
            );

            // Increment the index counter
            index++;
          });

          if (lastItem !== null) {
            this.nodesToExpand.push(lastItem);
          }
        }
      }

      // Else if a singular value, it checks to check it has childIdentityFields otherwise it should not be inserting the value
      else if (
        /* matrix.columns.root.hasOwnProperty('childIdentityFields') ||
        matrix.rows.root.hasOwnProperty('childIdentityFields') */
        Object.prototype.hasOwnProperty.call(matrix.columns.root, 'childIdentityFields') ||
        Object.prototype.hasOwnProperty.call(matrix.rows.root, 'childIdentityFields')
      ) {
        // Insert the value so it is not undefined or null
        rowObj[Object(columnDefs)[0]['field']] = row.value;
      }

      // Depending on the expansionUp check, we need to insert the rowObj in different ways
      if (!this.expandUpwards && this.rowChildrenNodes.length > 0) {
        // Clear out the rowData
        rowData.length = 0;

        // Create temporary holder array
        const tempHolderNodes = [];

        for (const node of this.rowChildrenNodes) {
          tempHolderNodes.push(node);
        }

        // Put the last item in tempHolderNodes first
        tempHolderNodes.unshift(tempHolderNodes.pop());

        this.rowChildrenNodesSorted = [
          ...this.rowChildrenNodesSorted,
          ...tempHolderNodes,
        ];

        this.rowChildrenNodes = [];

        for (const node of this.rowChildrenNodesSorted) {
          rowData.push(node);
        }
      }

      

      // If multiple level rows, clear out the rowData before inserting it. AND EXPAND UP!
      else if (this.rowChildrenNodes.length > 0) {
        // Clearing out the arrray
        rowData.length = 0;
        // Looping through the rowChildrenNodes and pushing them to rowData
        this.rowChildrenNodes.forEach((child) => {
          rowData.push(child);
        });
      }

      if (
        rowData.length !== this.rowChildrenNodesSorted.length &&
        this.rowChildrenNodesSorted.length > 0
      ) {
        rowData.pop();
      }

      // Push into rowData
      rowData.push(rowObj);
    });

    // console.log(rowData);

    // We fix the last row header of "Total" by getting the length of the array and changing the name on the last object. Other wise it will remain blank
    const lengthOfRowData = rowData.length;

    // Insert "Total" as last row header with a keyof typeof to ensure the string can access the object index. Checks first if undefined
    if (
      rowData[lengthOfRowData - 1][dynamicHeader as keyof typeof rowData] ===
      undefined
    ) {
      // Ensure a "Total does not show up for measures only matrix"
      if (
        /* matrix.columns.root.hasOwnProperty('childIdentityFields') ||
        matrix.rows.root.hasOwnProperty('childIdentityFields') */
        Object.prototype.hasOwnProperty.call(matrix.columns.root, 'childIdentityFields') ||
        Object.prototype.hasOwnProperty.call(matrix.rows.root, 'childIdentityFields')
      ) {
        //  Insert Total in the first column last row
        rowData[lengthOfRowData - 1][dynamicHeader as keyof typeof rowData] =
          'Total';
      }
    }

    // Update the object's rowData
    this.rowData = rowData;
  }


  private static finalizeGrid(columnDefs, rowData, dynamicHeader) {
    // let the grid know which columns and what data to use
    /*  let gridOptions = this.gridOptions; */
    const gridOptions = this.gridOptions;

    // Insert column width into grid options defaultColDef from the formatting settings columnCard
    gridOptions['defaultColDef']['width'] =
      this.formattingSettings.columnCard.columnWidth.value;

    // Check if wrapValues and wrapHeaders are enabled, otherwise set to false
    gridOptions['defaultColDef']['wrapText'] =
      this.formattingSettings.columnCard.wrapValues.value;
    gridOptions['defaultColDef']['wrapHeaderText'] =
      this.formattingSettings.columnCard.wrapHeaders.value;

    // Check if drag is enabled
    gridOptions['defaultColDef']['resizable'] =
      this.formattingSettings.columnCard.enableDrag.value;

    // Check if auto row height is enabled
    gridOptions['defaultColDef']['autoHeight'] =
      this.formattingSettings.columnCard.autoRowHeight.value;

    // Check if auto header height is enabled
    gridOptions['defaultColDef']['autoHeaderHeight'] =
      this.formattingSettings.columnCard.autoHeaderHeight.value;

    // Checking if there is a total row to pin to the bottom of the grid
    const total = rowData.pop();

    // If the total is not "Total" then push it back into the rowData
    if (total[dynamicHeader as keyof typeof total] !== 'Total') {
      rowData.push(total);
    }

    // If only measures and columns are added then push it back (As you only have one row and it is the total)
    if (rowData.length === 0) {
      rowData.push(total);
    }

    // A total row in the class field in case it should be refered to later
    this.pinnedTotalRow = total;
    
    // Populate the gridOptions
    gridOptions['columnDefs'] = columnDefs;
    gridOptions['rowData'] = rowData;

    // The gridDiv
    const gridDiv: HTMLElement = document.createElement('div');
    gridDiv.className = 'ag-theme-alpine';
    gridDiv.id = 'myGrid';

    // Set the visibility style to hidden to avoid flickering
    gridDiv.style.visibility = 'hidden';

    // Creates the final Grid
    new agGrid.Grid(gridDiv, gridOptions);

    // Set the pinned bottom row data
    if (total[dynamicHeader as keyof typeof total] === 'Total') {
      gridOptions.api.setPinnedBottomRowData([total]);
    }

    // Return a finished DIV to be attached
    return gridDiv;
  }

  /**
     * The core function for the creation of the matrix
     */
  private static formatMatrix(dataView: powerbi.DataView) {
    this.resetState();

    const { matrix, dynamicHeader, rowHeight, headerHeight } = this.extractMetadata(dataView);

    const { gridCellStyling, categoryCellStyling, rowHeadersCard } = this.applyGridStyling();

    const columnDefs = [];

    let colId = 0;

    this.handleSingularMeasureMatrix(matrix, dataView, columnDefs, colId, dynamicHeader, gridCellStyling, categoryCellStyling, rowHeadersCard);

    this.defineColumns(matrix, dataView, columnDefs, colId, dynamicHeader, gridCellStyling, categoryCellStyling, rowHeadersCard);

    const rowData: object[] = [];
    this.buildRowData(matrix, columnDefs, rowData, dynamicHeader);

    //const total = rowData.pop();
    return this.finalizeGrid(columnDefs, rowData, dynamicHeader);
  }
  
  /**
   * Recursive function to traverse child nodes
   */
  private static async traverseChildNodes(
    child,
    columnDefs,
    parentHeader,
    identityHolder
  ) {
    // Identity holder in a separate array to avoid reference duplication
    identityHolder = [...identityHolder];

    // Create the rowObj
    const rowObj = {};

    // Deconstruct row into values
    const rowValues = child.values;

    // Index iterator
    let index = 0;

    // Check if child.value (aka the header) is undefined, if so grab the parentHeader
    let rowHeader = child.value;

    if (child.value === undefined) {
      rowHeader = parentHeader;
    }

    // Insert row header
    rowObj[Object(columnDefs)[index]['field']] = String(rowHeader);

    // Add the nodes to the nodesToExpand array so expand buttons can be programatically added & for expansion identification
    if (
      /* child.hasOwnProperty('isCollapsed') ||
      child.hasOwnProperty('identity') */
      Object.prototype.hasOwnProperty.call(child, 'isCollapsed') ||
      Object.prototype.hasOwnProperty.call(child, 'identity')
    ) {
      this.nodesToExpand.push(child);
    }

    // If identity holder is not subTotal, push it
    if (child.isSubtotal !== true) {
      identityHolder.push(child);
    }

    // Push the identity holder into the rowObj
    rowObj['identity'] = identityHolder;

    // Increment index
    index++;

    // Valuesource index back up (For calculation groups)
    let valueSourcesIndexBackUp = 0;

    // Create a temporary holding array for sorting
    const tempRowChildren = [...this.tempRowChildren];

    // Try catch as sometimes it is values and sometimes it is children, but they never exist together in the dataView.
    try {
      // Loop through last level of children and insert into rowObj
      Object.values(rowValues).forEach((value) => {
        const valueFormatted = this.valueFormatterMatrix(
          value['value'],
          value['valueSourceIndex'] === undefined
            ? valueSourcesIndexBackUp
            : value['valueSourceIndex'],
          rowValues
        );

        rowObj[Object(columnDefs)[index]['field']] = valueFormatted;

        // Increment the index counter
        index++;
        valueSourcesIndexBackUp++;
      });

      // If expand upwards splice and merge the arrays otherwise just insert
      if (this.expandUpwards === false) {
        // Push into both arrays
        this.tempRowChildren.push(rowObj);
        this.rowChildrenNodes.push(rowObj);

        // Push into the local array
        tempRowChildren.push(rowObj);

        // Get the last item of the tempRowChildren array
        const LastITem = tempRowChildren.pop();

        // Insert it first
        tempRowChildren.unshift(LastITem);

        // Splice the list to avoid duplicates and then merge
        const lengthOfRowChildrenNodes = this.rowChildrenNodes.length;
        const lengthOfTempNodes = tempRowChildren.length;
        const difference = lengthOfRowChildrenNodes - lengthOfTempNodes;

        this.rowChildrenNodes.splice(difference, lengthOfTempNodes);

        this.rowChildrenNodes = [...this.rowChildrenNodes, ...tempRowChildren];
      } else {
        this.rowChildrenNodes.push(rowObj);
      }

      // If catch then it is children and we need to recursively loop through them
    } catch {
      // Contains the children in case expanse upward is false
      const tempRowChildrenContained = [...this.tempRowChildren];

      // If expand upwards is false, reset the tempRowChildren array
      if (this.expandUpwards === false) {
        this.tempRowChildren.length = 0;
      }

      let lastItem = null;
      // If the expandUp is true, pop the last item in the array

      if (this.formattingSettings.expansionCard.expandUp.value === true) {
        lastItem = this.nodesToExpand.pop();
      }

      // Loop through children and insert into rowObj
      child.children.forEach((child) => {
        this.traverseChildNodes(child, columnDefs, rowHeader, identityHolder);

        index++;
      });

      // Insert it after every other item
      if (lastItem !== null) {
        this.nodesToExpand.push(lastItem);
      }

      // If expand upwards, unshift and merge the temprow children and temprowChildrenContained arrays
      if (this.expandUpwards === false) {
        // tempRowChildrenContained.unshift(tempRowChildrenContained.pop());

        this.tempRowChildren.unshift(this.tempRowChildren.pop());

        // tempRowChildrenContained.unshift(this.tempRowChildren.pop());

        tempRowChildren.reverse();

        this.tempRowChildren = [
          ...tempRowChildrenContained,
          ...this.tempRowChildren,
        ];
      }
    }
  }

  /**
   * Handles the selection on clicks
   */
  private static async selectOnClick(cell) {
    // console.log('Cell was clicked', cell);

    // console.log(this.formattingSettings);

    // Make sure it is not propagating through the expand button (This is due us needing the grid API to be triggered for traversing algorithm information)
    if (cell.event.target.classList[0] === 'expandButton') {
      this.startExpansion(cell);
      return;
    }

    // multiSelect variable
    let multiSelect = false;

    // Extract the row and column from the cell
    // Selected row, assuming there are no levels in the dataView
    let selectedRow;

    // Check level length
    if (this.dataView.matrix.rows.levels.length > 1) {
      // Traverse through the rowChildren and try to find the value of the row
      const rowHeader = Object.values(cell.data)[0];

      // When found it becomes selectedRow
      selectedRow = this.traverseSelection(
        rowHeader,
        this.dataView.matrix.rows.root.children
      );
    }
    // Assuming no levels
    else {
      // SelectedRow per default value
      selectedRow = this.rowChildren[cell.rowIndex];
    }

    // Selected column
    const selectedColumn = this.columnChildren[cell.column.colId - 1];

    // console.log(selectedColumn);
    // console.log(selectedRow);

    // Create selectionId
    const selectionId = this.visualMapping(selectedRow, selectedColumn);

    // Check if right clicked for context menu
    if (cell.event.type === 'contextmenu') {
      this.selectionManager.showContextMenu(selectionId, {
        x: cell.event.clientX,
        y: cell.event.clientY,
      });

      return;
    }

    // Check if multiSelect is true
    if (cell.event.ctrlKey) {
      multiSelect = true;
    }

    // Get the selected class and turn it off
    const selected = document.querySelector('.selected');

    // Get the cell via the event path and color it to show it is selected
    let cellElement = cell.eventPath[0] as HTMLElement;

    // Check if cell contains class "gridCell" or "categoryCell"
    if (
      !cellElement.classList.contains('gridCell') &&
      !cellElement.classList.contains('categoryCell')
    ) {
      cellElement = cell.eventPath[2] as HTMLElement;
    }

    // Add the selected class if class does not contain selected and remove the previously selected
    if (!cellElement.classList.contains('selected')) {
      cellElement.classList.add('selected');
      if (selected) {
        selected.classList.remove('selected');
      }
    } else {
      cellElement.classList.remove('selected');
    }

    // Apply the selection
    this.selectionManager.select(selectionId, multiSelect);
  }

  /**
   * Traversing the selection
   */
  private static traverseSelection(rowHeader, children) {
    // A sentinel value to stop loops
    let sentinelValueStop = false;

    // The variable that is to be returned
    let childToBeReturned;

    // Loop through children until you find the rowheader
    children.forEach((child) => {
      // Abort the loop if child is fiund
      if (sentinelValueStop) {
        return;
      }

      // Look for the right child
      if (child.value === rowHeader) {
        // If found, set the sentinel value and set the child to the functional scope variable
        sentinelValueStop = true;
        childToBeReturned = child;

        // Finish the iteration
        return childToBeReturned;
      }

      // Recursively go through the tree if child has not been found and it has children
      if (child.children && !sentinelValueStop) {
        // Set the variable
        childToBeReturned = this.traverseSelection(rowHeader, child.children);

        // If the child is not undefined / null, set the sentinelvalue to stop further iterations
        if (childToBeReturned) {
          sentinelValueStop = true;
        }
        // Finish the iteration
        return childToBeReturned;
      }
    });

    // Return the child through the recursive chain
    return childToBeReturned;
  }

  /**
   * Adds the body event listeners
   */
  private static addBodyEventListeners() {
    // Get the grid body
    const gridBody = document.getElementsByClassName('ag-body-viewport');

    // Add event listeners for selection to deselect
    gridBody[0].addEventListener('click', (e) => {
      const event = e as PointerEvent;

      // See if the target is the body and not a cell, and cast it so ti works
      const target = (event.target as Element).classList[0];

      // If true, clear the selection
      if (target === 'ag-body-viewport') {
        this.selectionManager.clear();
      }
    });

    // Event listener for the context menu (right click)
    gridBody[0].addEventListener('contextmenu', (e) => {
      // Dummy selection id as we are not selecting anything.
      const selectionId = {};

      // Convet event to MouseEvent
      const event = e as MouseEvent;

      // Bring up context menu via Power BI API
      this.selectionManager.showContextMenu(selectionId, {
        x: event.clientX,
        y: event.clientY,
      });
    });
  }

  /**
   * Event listeners for the headers
   */
  private static async addHeaderEventListeners() {
    // Get the header elements (Important to await!)
    const headerElements = await document.querySelectorAll(
      '.ag-header-cell-comp-wrapper'
    );

    // Create a sentinel value for the forEach loop in the case of all measures
    let allMeasuresListened = false;
    

    // Loop through the header elements
    headerElements.forEach((element, index) => {
      // Convert the type of the element to something that can be used
      /* let elementUnknown = element as unknown; */
      const elementUnknown = element as unknown;
      const elementHTML = elementUnknown as HTMLElement;

      // Add the context menu to every header
      this.addContextMenuEventListener(elementHTML);

      // Add special eventListener to the first column (For sorting)
      if (index === 0) {
        let dataSource;
        // TRY to Get the row query name to sort
        try {
          dataSource = this.dataView.matrix.rows.levels[0].sources[0];
        } catch {
          // Else it is a column
          dataSource = this.dataView.matrix.columns.levels[0].sources[0];
        }

        // Adds the sorting event listener
        this.addSortingEventlistener(elementHTML, dataSource);

        // Return to finish this iteration of the loop
        return;
      }

      // If there are multiple measures and no columns, proceed to add a sorting event listener for every column.
      // We are assuming no columns has a length of 1

      if (
        this.dataView.matrix.valueSources.length > 1 ||
        this.dataView.matrix.columns.root.children.length == 1
      ) {
        // Put value sources in, -1 as the index will be n+1 by the time it comes here due to returning from the first column
        const dataSource = this.dataView.matrix.valueSources[index - 1];

        // Loop through the add event listener
        this.addSortingEventlistener(elementHTML, dataSource);

        // Set sentinel value to true
        allMeasuresListened = true;

        // Return to quit this iteration of the loop
        return;
      }

      // Check the last header if it is "Total" and implement custom sorting on it
      if (index === headerElements.length - 1) {
        // Get the textContent of the header table
        const headerLabel = elementHTML.getElementsByClassName(
          'ag-header-cell-text'
        )[0].textContent;

        // Check if it is "Total"
        if (headerLabel === 'Total' || headerLabel.includes('Invalid')) {
          // Get the row query name to sort
          const dataSource = this.dataView.matrix.valueSources[0];

          //  Add the event listener
          this.addSortingEventlistener(elementHTML, dataSource);

          // Return to end this it iteration of the loop
          return;
        }
      }

      // Add the selection to every header (This is last as not to have duplicate event listeners)
      this.addSelectionEventListener(elementHTML, index);
    });
  }

  // Add a selection header event listener
  private static addSelectionEventListener(element, index) {
    // Add the event listener for selections on non-sorting columns
    element.addEventListener('click', (e) => {
      // Get column
      const selectedColumn = this.columnChildren[index - 1];

      // Create selection ID (Empty object in row due to us selecting only a column)
      const selectionId = this.visualMapping({}, selectedColumn);

      // Enabling multiselect with false as default
      let multiSelect = false;

      // Check if multiSelect is true
      if (e.ctrlKey) {
        multiSelect = true;
      }
      // Creating selection
      this.selectionManager.select(selectionId, multiSelect);
    });
  }

  /**
   * Adds the context menu event listeners
   */
  private static addContextMenuEventListener(element) {
    // Adding the context menu on every column header
    element.addEventListener('contextmenu', (e) => {
      e.preventDefault();

      // Dummy selection id as we are not selecting anything.
      const selectionId = {};

      this.selectionManager.showContextMenu(selectionId, {
        x: e.clientX,
        y: e.clientY,
      });
    });
  }

  /**
   * Adds the sorting event listeners
   */
  private static addSortingEventlistener(element, source) {
    element.addEventListener('click', (e) => {
      // Get the row query name to sort
      const sortingQueryName = source.queryName;

      // Get the sorting type and inverse it
      const sortOrder = source.sort;

      // Set the sorting
      const sortDirection =
        sortOrder === 2
          ? powerbi.SortDirection.Ascending
          : powerbi.SortDirection.Descending;

      // The sorting arguments, assuming Descending as default
      /* let sortArgs = */ 
      const sortArgs = {
        sortDescriptors: [
          {
            queryName: sortingQueryName,
            sortDirection: sortDirection,
          },
        ],
      };

      // Apply the sorting
      this.host.applyCustomSort(sortArgs);
    });
  }

  // Clears the grid in preparation for a new set of data
  public static clearGrid() {
    this.gridOptions.api.setRowData([]);
    this.gridOptions.api.setColumnDefs([]);
    this.gridOptions.api.hideOverlay();
  }

  // Maps the node to a selectionId
  private static visualMapping(row, column) {
    // Creates the selection id
    const nodeSelection = this.host
      .createSelectionIdBuilder()
      .withMatrixNode(column, this.columnLevels)
      .withMatrixNode(row, this.rowLevels)
      .createSelectionId();

    // Returns the selection
    return nodeSelection;
  }

  /**
   * This functions creates the mapping for the row expansion
   */
  private static rowMapping(nodeLineage) {
    // Create the selectionbuilder
    const nodeSelectionBuilder = this.host.createSelectionIdBuilder();

    // Creates the selection id via looping through parents + current node
    for (const node of nodeLineage) {
      nodeSelectionBuilder.withMatrixNode(node, this.rowLevels);
    }

    // Create the selectionID
    const nodeSelectionId = nodeSelectionBuilder.createSelectionId();

    return nodeSelectionId;
  }

  /**
   *  This adds the expand buttons for every row
   */
  private static AddExpandButtons() {
    // Check if addButtons is in the formattingSetting card otherwise return
    if (this.formattingSettings.expansionCard.enableButtons.value === false) {
      return;
    }

    // Get the rows via the ag-row class
    const rows = document.querySelectorAll('.ag-row');

    // Loop through the rows
    for (const rowHeader of Object(rows)) {
      const rowHeaderTextcontent = rowHeader.children[0].textContent;
      const rowHeaderDIV = rowHeader.children[0];

      // For every rowHeader, loop through the nodesToExpand Array to find a match
      for (const node of this.nodesToExpand) {
        // Check if there is an equal value
        if (rowHeaderTextcontent === node.value) {
          // Check that there is no button already added
          const potentialButtons =
            rowHeaderDIV.querySelectorAll('.expandButton');

          // If length longer than 0, skip to next loop
          if (potentialButtons.length != 0) {
            continue;
          }

          // Check that the level can in fact be expanded
          const levels = this.dataView.matrix.rows.levels;

          const nodeLevel = node.level;

          const nodeLevelInTree = levels[nodeLevel];

          if (
            nodeLevelInTree['canBeExpanded'] === false ||
            !Object.prototype.hasOwnProperty.call(nodeLevelInTree, 'canBeExpanded')
            /* !nodeLevelInTree.hasOwnProperty('canBeExpanded') */
          ) {
            continue;
          }

          // Create the button and style it
          const button = document.createElement('button');

          // CSS STYLING
          button.classList.add('expandButton');
          button.textContent = '+';
          button.style.left = `${(nodeLevel + 1) * 10}px`;

          // Append to the div
          rowHeaderDIV.insertBefore(button, rowHeaderDIV.firstChild);
        }
      }
    }
  }

  /**
   *  This function enables the button to expand up and down
   */
  private static startExpansion(cell) {
    // Get the selectedRow

    // Grab the lineage of the selected row
    const selectedRowIdentityLineage = cell.data.identity;

    // Create a selectionID
    const selectionID = this.rowMapping(selectedRowIdentityLineage);

    this.selectionManager.toggleExpandCollapse(selectionID);
  }

  /**
   * This functions formats the expanded rows.
   */
  private static formatExpandedRows() {
    // console.log(this.formattingSettings);

    // IF checks to make sure function is not running needlessly
    if (
      /* !this.dataView.matrix.columns.root.hasOwnProperty(
        'childIdentityFields'
      ) &&
      !this.dataView.matrix.rows.root.hasOwnProperty('childIdentityFields')
     */
    !Object.prototype.hasOwnProperty.call(this.dataView.matrix.columns.root, 'childIdentityFields') &&
    !Object.prototype.hasOwnProperty.call(this.dataView.matrix.rows.root, 'childIdentityFields')
      ) {
      return;
    }

    try {
      // If the row level length is 1, return as we do not want to format rows that cannot be expanded/Have buttons
      if (
        /* !this.dataView.matrix.rows.levels[0].hasOwnProperty('canBeExpanded') */
        !Object.prototype.hasOwnProperty.call(this.dataView.matrix.rows.levels[0], 'canBeExpanded')
      ) {
        return;
      }
    } catch {
      return;
    }

    // Get the expansionCard
    const expansionCard = this.formattingSettings.expansionCard;

    // Get all the rows
    const rows = document.querySelectorAll('.ag-row');

    // Loop through them
    for (const rowHeader of Object(rows)) {
      // If singular measure, do not run this function as there is no expansion

      // Get the row-id from the HTML attribute
      const rowId = rowHeader.getAttribute('row-id');

      // Get the corresponding node from the nodesToExpand Array
      const rowNodeData = this.nodesToExpand[rowId];

      // The rowHeaderTextContent to skip "Total" to avoid an error but also indent it
      const rowHeaderTextcontent = rowHeader.children[0].textContent;

      // The rowHeaderDIV for CSS transformation
      const rowHeaderDIV = rowHeader.children[0];

      // INDENTATION
      // Get the value from the formatting settings
      let indentation =
        this.formattingSettings.expansionCard.indentationValue.value;

      // Try catch in case of unforseen errors (Should not be any but for good measure) Also indent it to the same as first level nodes
      try {
        if (rowHeaderTextcontent === 'Total') {
          // Check if indentation is enabled
          if (
            this.formattingSettings.expansionCard.enableIndentation.value ===
            true
          ) {
            // Check if buttons are enabled
            if (
              this.formattingSettings.expansionCard.enableButtons.value === true
            ) {
              rowHeaderDIV.style.textIndent = `${indentation}px`;
              rowHeaderDIV.style.textAlign = 'start';
            }

            continue;
          }
        }

        // Ensure node level is 1 for multiplication
        let nodeLevel = rowNodeData.level === 0 ? 1 : rowNodeData.level + 1;

        // If buttons are disabled, remove one level
        if (
          this.formattingSettings.expansionCard.enableButtons.value === false
        ) {
          nodeLevel = nodeLevel - 1;
        }

        indentation = parseInt(nodeLevel) * indentation;

        // Check if indentation is enabled
        if (
          this.formattingSettings.expansionCard.enableIndentation.value === true
        ) {
          // Put in the indentation
          rowHeaderDIV.style.textIndent = `${indentation}px`;
        }

        // Bold rows that have been expanded and change the button text to "-"
        if (rowNodeData.isCollapsed === false) {
          // Set the background color via image left right gradient

          // To be refered to later in the expandedRows function
          rowHeader.classList.add('expandedRow');

          // Set alignment
          for (const child of rowHeader.children) {
            // Set the bold
            child.style.fontWeight = expansionCard.enableBold.value
              ? 'bold'
              : 'normal';

            // Set the italic
            child.style.fontStyle = expansionCard.enableItalic.value
              ? 'italic'
              : 'normal';

            // child.style.backgroundColor = 'red';
            child.style.justifyContent = expansionCard.alignment.value.value;

            // Set the fontFamily
            child.style.fontFamily = expansionCard.fontFamily.value;

            // set the fontSize
            child.style.fontSize = `${expansionCard.fontSize.value}px`;

            // Set the fontColor
            child.style.color = expansionCard.fontColor.value.value;
          }

          // Get the value from the formatting settings
          const borderWidth = expansionCard.borderWidth.value;

          // Get the border color from the formatting settings
          const borderColor = hex2rgba(
            expansionCard.borderColor.value.value,
            expansionCard.borderOpacity.value
          );

          // Get the border style from the formatting settings
          const borderStyle = expansionCard.borderStyle.value.value;

          // Set the style TOP
          rowHeader.style.borderTop = expansionCard.enableTopBorder.value
            ? `${borderWidth}px ${borderStyle} ${borderColor}`
            : 'none';

          // Set the style BOTTOM
          rowHeader.style.borderBottom = expansionCard.enableBottomBorder.value
            ? `${borderWidth}px ${borderStyle} ${borderColor}`
            : 'none';

          this.gridOptions.api
            .getRowNode(rowId)
            .setRowHeight(expansionCard.height.value);

          // Rowheader style via background image
          rowHeader.style.backgroundImage = `linear-gradient(to right, ${hex2rgba(
            expansionCard.backgroundColor.value.value,
            expansionCard.opacity.value
          )} 100%, white 100%)`;

          // Change the button text to "-"
          const button = rowHeaderDIV.querySelector('.expandButton');
          button.textContent = '-';
        }
      } catch (e) {}
    }
  }

  /**
   * Format the rows to get a filled background color. The function was reduced from a more complex function as the rest was integrated into the grid
   */
  private static formatRows() {
    // Get all the rows
    const rows = document.querySelectorAll('.ag-row');

    const rowCard = this.formattingSettings.rowCard;

    // Get the background color
    const backgroundColor = hex2rgba(
      rowCard.backgroundColor.value.value,
      rowCard.opacity.value
    );

    for (const row of Object(rows)) {
      // Apply background colors
      // (row.style.backgroundColor = backgroundColor),

      row.style.backgroundImage = `linear-gradient(to right, ${backgroundColor} 100%, white 100%)`;

      // Row borders
      (row.style.borderTop = rowCard.enableTopBorder.value
        ? `${rowCard.borderWidth.value}px ${
            rowCard.borderStyle.value.value
          } ${hex2rgba(
            rowCard.borderColor.value.value,
            rowCard.borderOpacity.value
          )}`
        : 'none'),
        (row.style.borderBottom = rowCard.enableBottomBorder.value
          ? `${rowCard.borderWidth.value}px ${
              rowCard.borderStyle.value.value
            } ${hex2rgba(
              rowCard.borderColor.value.value,
              rowCard.borderOpacity.value
            )}`
          : 'none');
    }
  }

  /**
   * Format the row headers
   */
  private static formatRowHeaders() {
    // Check that rowHeaders card is enabled
    if (this.formattingSettings.rowHeadersCard.enableCard.value === false) {
      return;
    }

    // Get all the rows
    const rows = document.querySelectorAll('.ag-row');

    // The rowheaders card
    const rowHeadersCard = this.formattingSettings.rowHeadersCard;

    // Get the row card
    const rowCard = this.formattingSettings.rowCard;

    // Get the rowHeaders background color & opacity
    const backgroundColor = hex2rgba(
      rowHeadersCard.backgroundColor.value.value,
      rowHeadersCard.opacity.value
    );

    for (const row of Object(rows)) {
      // Dynamic rowBackgroundColor that depends on if a row is expanded or not
      let rowBackgroundColor;

      // Get the row background color and opacity
      rowBackgroundColor = hex2rgba(
        rowCard.backgroundColor.value.value,
        rowCard.opacity.value
      );
      // Get the width of a categoryCell
      const width = row.children[0].offsetWidth;

      if (row.classList.contains('expandedRow')) {
        rowBackgroundColor = hex2rgba(
          this.formattingSettings.expansionCard.backgroundColor.value.value,
          this.formattingSettings.expansionCard.opacity.value
        );

        // Replace header styles for expanded Rows
        row.children[0].style.fontFamily = rowHeadersCard.fontFamily.value;
        row.children[0].style.fontSize = `${rowHeadersCard.fontSize.value}px`;
        row.children[0].style.color = rowHeadersCard.fontColor.value.value;
        row.children[0].style.fontWeight = rowHeadersCard.enableBold.value
          ? 'bold'
          : 'normal';
        row.children[0].style.fontStyle = rowHeadersCard.enableItalic.value
          ? 'italic'
          : 'normal';

        row.children[0].style.justifyContent =
          rowHeadersCard.alignment.value.value;
      }

      // Set a linear gradient with the header backgroundColor and the row backgroundColor with a cut off where the width of the header is
      row.style.backgroundImage = `linear-gradient(to right, ${backgroundColor} ${width}px, ${rowBackgroundColor} ${width}px)`;
    }
  }

  /**
   * Format the columns
   */
  private static formatColumns(gridApi) {
    // Disable persisted properties

    // Check if auto width is enabled
    if (this.formattingSettings.columnCard.enableAutoWidth.value === true) {
      gridApi.columnApi.autoSizeAllColumns();
    }
  }

  /**
   * THe value formatter for the matrix
   */
  private static valueFormatterMatrix(value, valueSourceIndex, row) {
    // Get the value sources from the dataView
    const valueSources = this.dataView.matrix.valueSources;

    // For some reason, a value source index is always provided, except in the cases of the first index. Therefore if undefined, make it a 0
    if (valueSourceIndex === undefined) {
      valueSourceIndex = 0;
    }

    // Declare the variables for use later
    let formatString;
    let valueSource;

    // Get length of valueSources to ensure we not go out of bounds
    const length = valueSources.length;

    // Ensure that the valueSourceIndex is not out of bounds
    if (valueSourceIndex <= length - 1) {
      // Get the valueSource
      valueSource = valueSources[valueSourceIndex];
    }
    // Backup incase if it fails with calc group columns
    if (length === 1) {
      valueSource = valueSources[0];
    }

    // Get the format string
    formatString = valueSource.format;

    try {
      // If the attempt fails (as in undefined) then it means the project is using dynamic font strings and must search in different place
      if (formatString === undefined) {
        formatString = row[valueSourceIndex].objects.general.formatString;
      }
    } catch {
      // If the formatString is undefined OR NULL (Which happens if an implicit measure OR the measure is set to automatic) then set to the default formatString
      if (
        formatString === undefined ||
        formatString === null ||
        formatString === ''
      ) {
        formatString = '0.0';
      }
    }
    // Ensure the formatString has a format
    if (
      formatString === undefined ||
      formatString === null ||
      formatString === ''
    ) {
      formatString = '0.0';
    }

    // Create the value formatter with the formatString
    /* let iValueFormatter = valueFormatter.create({ format: formatString }); */
    const iValueFormatter = valueFormatter.create({ format: formatString });

    // Format the value
    let formattedValue = iValueFormatter.format(value);

    // We want blanks to be invisible rather than written out as blanks
    if (formattedValue === '(Blank)') {
      formattedValue = '';
    }

    // Return the formattedValue
    return formattedValue;
  }

  /**
   * Format the column headers
   */
  private static formatColHeaders() {
    // Check that rowHeaders card is enabled
    if (this.formattingSettings.colHeadersCard.enableCard.value === false) {
      return;
    }

    // Get all the rows
    const rows = document.querySelectorAll('.ag-header-cell');

    // Get the row font
    const font = this.formattingSettings.colHeadersCard.fontFamily.value;

    // Get the row font size
    const fontSize = this.formattingSettings.colHeadersCard.fontSize.value;

    // Get the row font color
    const fontColor =
      this.formattingSettings.colHeadersCard.fontColor.value.value;

    // Get the bold value
    const bold = this.formattingSettings.colHeadersCard.enableBold.value;

    // Get the italic value
    const italic = this.formattingSettings.colHeadersCard.enableItalic.value;

    // Get the background color
    const backgroundColor =
      this.formattingSettings.colHeadersCard.backgroundColor.value.value;

    // Get the border width
    const borderWidth =
      this.formattingSettings.colHeadersCard.borderWidth.value;

    // Get the border color
    const borderColor = hex2rgba(
      this.formattingSettings.colHeadersCard.borderColor.value.value,
      this.formattingSettings.colHeadersCard.borderOpacity.value
    );

    // Get the border style
    const borderStyle =
      this.formattingSettings.colHeadersCard.borderStyle.value.value;

    // Get the Top border enabled
    const topBorder =
      this.formattingSettings.colHeadersCard.enableTopBorder.value;

    // Get the right border enabled
    const rightBorder =
      this.formattingSettings.colHeadersCard.enableRightBorder.value;

    // Get the left border
    const leftBorder =
      this.formattingSettings.colHeadersCard.enableLeftBorder.value;

    // Get alignment
    const alignment =
      this.formattingSettings.colHeadersCard.alignment.value.value;

    // Get the Bottom border enabled
    const bottomBorder =
      this.formattingSettings.colHeadersCard.enableBottomBorder.value;

    for (const rowContainer of Object(rows)) {
      const row = rowContainer.querySelector('.ag-header-cell-label');

      // Set the background Color to the parent element of the row to avoid gaps in the background color (Perhaps changed later to cells, depending on specific columns)
      rowContainer.parentElement.style.backgroundColor = hex2rgba(
        backgroundColor,
        this.formattingSettings.colHeadersCard.opacity.value
      );

      const rowContainerParent = rowContainer.parentElement
        .parentElement as HTMLElement;

      // Set the row font
      row.style.fontFamily = font;

      // Set the row font size
      row.style.fontSize = `${fontSize}px`;

      // Set the row font color
      row.style.color = fontColor;

      // Set the bold
      row.style.fontWeight = bold ? 'bold' : 'normal';

      // Set the italic
      row.style.fontStyle = italic ? 'italic' : 'normal';

      // Set the top border
      rowContainerParent.style.borderTop = topBorder
        ? `${borderWidth}px ${borderStyle} ${borderColor}`
        : 'none';

      // Set the bottom border
      rowContainerParent.style.borderBottom = bottomBorder
        ? `${borderWidth}px ${borderStyle} ${borderColor}`
        : 'none';

      // Set the right border
      rowContainer.style.borderRight = rightBorder
        ? `${borderWidth}px ${borderStyle} ${borderColor}`
        : 'none';

      // Set the left border
      rowContainer.style.borderLeft = leftBorder
        ? `${borderWidth}px ${borderStyle} ${borderColor}`
        : 'none';

      // Set alignment
      for (const child of row.children) {
        child.style.justifyContent = alignment;
      }
    }
  }
  /**
   * Format specific rows
   */
  private static formatSpecificRows() {
    // Get all the categoryCells
    const categoryCells = document.querySelectorAll('.categoryCell');

    // If categoryCells is null, return
    if (categoryCells === null) {
      return;
    }

    // Loop through the formatting settings cards array
    for (const card of this.formattingSettings.cards) {
      // Check if the card contains the string "specificRow" via the name attribute
      const name = card.name;

      if (!name.includes('specificRow')) {
        continue;
      }

      // Check if the enableCard is enabled
      if (card.enableCard.value === false) {
        continue;
      }

      if (card.savedName.value === '') {
        continue;
      }

      // Get the applicable rows
      const applicableRows = card.savedName.value.split(',');

      // Get the rowHeaderFontFamily
      const rowHeaderFontFamily = card.rowHeaderFontFamily.value;

      // Get the height
      const height = card.height.value;

      // Get the RowHeaderFontColor
      const rowHeaderFontColor = card.rowHeaderFontColor.value.value;

      // Get the RowHeaderFontSize
      const rowHeaderFontSize = card.rowHeaderFontSize.value;

      // Get the RowHeaderBold
      const rowHeaderBold = card.rowHeaderBold.value;

      // Get the RowHeaderItalic
      const rowHeaderItalic = card.rowHeaderItalic.value;

      // Get the RowHeaderAlignment
      const rowHeaderAlignment = card.rowHeaderAlignment.value.value;

      // Get the RowHeaderBackground
      const rowHeaderBackground = card.rowHeaderBackground.value.value;

      // Get the row font
      const font = card.fontFamily.value;

      // Get the row font size
      const fontSize = card.fontSize.value;

      // Get the row font color
      const fontColor = card.fontColor.value.value;

      // Get the bold value
      const bold = card.enableBold.value;

      // Get the italic value
      const italic = card.enableItalic.value;

      // Get the background color
      const backgroundColor = card.backgroundColor.value.value;

      // Get the border width
      const borderWidth = card.borderWidth.value;

      // Get the border color
      const borderColor = hex2rgba(
        card.borderColor.value.value,
        card.borderOpacity.value
      );

      // Get the border style
      const borderStyle = card.borderStyle.value.value;

      // Get the Top border enabled
      const topBorder = card.enableTopBorder.value;

      // Get alignment
      const alignment = card.alignment.value.value;

      // Get the Bottom border enabled
      const bottomBorder = card.enableBottomBorder.value;

      // Loop through the applicale rows
      for (const displayName of applicableRows) {
        // A check to ensure the Total row is not formatted
        if (
          displayName === 'Total' ||
          displayName === this.formattingSettings.totalCard.savedName.value
        ) {
          continue;
        }
        // Loop through the category cells to find the correct row
        for (const rowHeader of Object(categoryCells)) {
          // Remove "+" or "-" from the row textContent
          const rowTextContent = rowHeader.textContent
            .replace('+', '')
            .replace('-', '');

          // Skip if they do not match
          if (
            rowTextContent !== displayName.replace('+', '').replace('-', '')
          ) {
            continue;
          }

          // Get the parent element to style the row
          const row = rowHeader.parentElement;

          // Get the row-id via query selector
          const rowId = row.getAttribute('row-id');

          // Get the row node
          const rowNode = this.gridOptions.api.getRowNode(rowId);

          // Ensure rowNode is not undefined otherwise the application fails
          if (rowNode != undefined) {
            // Set the row height via the grid API
            rowNode.setRowHeight(height);
          }

          // Set the top border
          row.style.borderTop = topBorder
            ? `${borderWidth}px ${borderStyle} ${borderColor}`
            : 'none';

          // Set the bottom border
          row.style.borderBottom = bottomBorder
            ? `${borderWidth}px ${borderStyle} ${borderColor}`
            : 'none';

          // Set child specific values
          for (const child of row.children) {
            child.style.justifyContent = alignment;

            // Set the row font
            child.style.fontFamily = font;

            // Set the row font size
            child.style.fontSize = `${fontSize}px`;

            // Set the row font color
            child.style.color = fontColor;

            // Set the bold
            child.style.fontWeight = bold ? 'bold' : 'normal';

            // Set the italic
            child.style.fontStyle = italic ? 'italic' : 'normal';

            // Set the indentation
            child.style.textIndent = `${card.indentation.value}px`;

            // This is set to remove the specific column background
            child.style.backgroundColor = 'RGBA(0,0,0,0)';
          }

          row.style.marginBottom = '5px';

          // Get the rowheader width for the linear gradient
          const rowHeaderWidth = rowHeader.offsetWidth;

          // Create a linear gradient where the color includes the headerColor and the row color where it cuts off at the header width
          row.style.backgroundImage = `linear-gradient(to right, ${hex2rgba(
            rowHeaderBackground,
            card.headerOpacity.value
          )} ${rowHeaderWidth}px, ${hex2rgba(
            backgroundColor,
            card.opacity.value
          )} ${rowHeaderWidth}px)`;

          // Set the rowHeader values to the rowHeader
          rowHeader.style.fontFamily = rowHeaderFontFamily;
          rowHeader.style.fontSize = `${rowHeaderFontSize}px`;
          rowHeader.style.color = rowHeaderFontColor;
          rowHeader.style.justifyContent = rowHeaderAlignment;
          rowHeader.style.fontWeight = rowHeaderBold ? 'bold' : 'normal';
          rowHeader.style.fontStyle = rowHeaderItalic ? 'italic' : 'normal';
          rowHeader.style.textIndent = `${card.rowHeaderIndentation.value}px`;
        }
      }
    }
  }

  /**
   * Format specific columns
   */
  private static formatSpecificColumns() {
    const headerCells = Object(
      document.querySelectorAll('.ag-header-cell-text')
    );

    // Loop through the formatting settings cards array
    for (const card of this.formattingSettings.cards) {
      // Check if the card contains the string "specificRow" via the name attribute
      const name = card.name;

      if (!name.includes('specificColumn')) {
        continue;
      }

      // Check if the enableCard is enabled
      if (card.enableCard.value === false) {
        continue;
      }

      // Get the columnAlignment from the card
      const alignment = card.columnAlignment.value.value;

      // Get the columncBackgroundColor
      const backgroundColor = hex2rgba(
        card.columnBackgroundColor.value.value,
        card.valuesOpacity.value
      );

      // Get the columnBold
      const bold = card.columnBold.value;

      // Get the columnItalic
      const italic = card.columnItalic.value;

      // Get the columnFontColor
      const fontColor = card.columnFontColor.value.value;

      // Get the columnFontFamily
      const font = card.columnFontFamily.value;

      // Get the columnFontSize
      const fontSize = card.columnFontSize.value;

      // Get the columnHeaderAlignment
      const columnHeaderAlignment = card.columnHeaderAlignment.value.value;

      // Get the columnHeaderBackgroundColor
      const columnHeaderBackgroundColor = hex2rgba(
        card.columnHeaderBackgroundColor.value.value,
        card.opacity.value
      );

      // Get the columnHeaderBold
      const columnHeaderBold = card.columnHeaderBold.value;

      // Get the columnHeaderFontColor
      const columnHeaderFontColor = card.columnHeaderFontColor.value.value;

      // Get the columnHeaderFontFamily
      const columnHeaderFontFamily = card.columnHeaderFontFamily.value;

      // Get the columnHeaderFontSize
      const columnHeaderFontSize = card.columnHeaderFontSize.value;

      // Get the columnHeaderItalic
      const columnHeaderItalic = card.columnHeaderItalic.value;

      // APPLIES TO HEADER AND REST OF COLUMN
      // Get the columnWidth
      const columnWidth = card.columnWidth.value;

      // Get the enableLeftBorder
      const enableLeftBorder = card.enableLeftBorder.value;

      // Get the enableRightBorder
      const enableRightBorder = card.enableRightBorder.value;

      // Get the borderColor
      const borderColor = hex2rgba(
        card.borderColor.value.value,
        card.borderOpacity.value
      );

      // Get the borderStyle
      const borderStyle = card.borderStyle.value.value;

      // Get the borderWidth
      const borderWidth = card.borderWidth.value;

      // Applicable columns via the card savedName
      const applicableColumns = card.savedName.value.split(',');

      for (const displayName of applicableColumns) {
        // Loop through the header cells to find the correct row
        for (const header of headerCells) {
          // Skip if they do not match
          if (header.textContent !== displayName) {
            continue;
          }

          // Get the parent of the parent of parent of parent of the header to style the column
          // This is due to line breaks being added from the top parent that could not get removed
          const headerParent =
            header.parentElement.parentElement.parentElement.parentElement;

          // Get the col-id attribute value form the headerParent
          const colId = headerParent.getAttribute('col-id');

          // Get the column via col-id from querySelectorAll
          const columnChildren = Object(
            document.querySelectorAll(`[col-id="${colId}"]`)
          );

          for (const child of columnChildren) {
            // // Apply styles to children
            child.style.justifyContent = alignment;
            child.style.backgroundColor = backgroundColor;
            child.style.fontFamily = font;
            child.style.fontSize = `${fontSize}px`;
            child.style.color = fontColor;
            child.style.fontWeight = bold ? 'bold' : 'normal';
            child.style.fontStyle = italic ? 'italic' : 'normal';
            child.style.borderLeft = enableLeftBorder
              ? `${borderWidth}px ${borderStyle} ${borderColor}`
              : 'none';

            child.style.borderRight = enableRightBorder
              ? `${borderWidth}px ${borderStyle} ${borderColor}`
              : 'none';

            child.style.borderBottom = 'none';
            child.style.borderTop = 'none';
          }

          // Apply to the headerParent and the header from the header variables
          // Set the column font
          header.style.fontFamily = columnHeaderFontFamily;
          header.style.fontStyle = columnHeaderItalic ? 'italic' : 'normal';
          header.style.fontWeight = columnHeaderBold ? 'bold' : 'normal';
          header.style.fontSize = `${columnHeaderFontSize}px`;
          header.style.color = columnHeaderFontColor;
          headerParent.style.backgroundColor = columnHeaderBackgroundColor;
          headerParent.style.display = 'flex';
          header.style.justifyContent = columnHeaderAlignment;

          // Apply the column width
          this.gridOptions.columnApi.setColumnWidths([
            { key: colId, newWidth: columnWidth },
          ]);
        }
      }
    }
  }

  /**
   * Format the Total Row
   */
  private static formatTotal() {
    // Get all the categoryCells
    const totalRow = document.querySelector('.ag-row-pinned') as HTMLElement;

    // Return if totalrow does not exist
    if (totalRow === null) {
      return;
    }

    // Formatting card
    const card = this.formattingSettings.totalCard;

    // Check if the enableCard is enabled
    if (card.enableCard.value === false) {
      totalRow.parentElement.parentElement.parentElement.style.display = 'none';
      return;
    }

    const totalTextContent =
      card.savedName === undefined ? 'Total' : card.savedName.value;

    // Get the rowHeaderFontFamily
    const rowHeaderFontFamily = card.rowHeaderFontFamily.value;

    // Get the height
    const height = card.height.value;

    // Get the RowHeaderFontColor
    const rowHeaderFontColor = card.rowHeaderFontColor.value.value;

    // Get the RowHeaderFontSize
    const rowHeaderFontSize = card.rowHeaderFontSize.value;

    // Get the RowHeaderBold
    const rowHeaderBold = card.rowHeaderBold.value;

    // Get the RowHeaderItalic
    const rowHeaderItalic = card.rowHeaderItalic.value;

    // Get the RowHeaderAlignment
    const rowHeaderAlignment = card.rowHeaderAlignment.value.value;

    // Get the RowHeaderBackground
    const rowHeaderBackground = card.rowHeaderBackground.value.value;

    // Get the row font
    const font = card.fontFamily.value;

    // Get the row font size
    const fontSize = card.fontSize.value;

    // Get the row font color
    const fontColor = card.fontColor.value.value;

    // Get the bold value
    const bold = card.enableBold.value;

    // Get the italic value
    const italic = card.enableItalic.value;

    // Get the background color
    const backgroundColor = card.backgroundColor.value.value;

    // Get the border width
    const borderWidth = card.borderWidth.value;

    // Get the border color
    const borderColor = hex2rgba(
      card.borderColor.value.value,
      card.borderOpacity.value
    );

    // Get the border style
    const borderStyle = card.borderStyle.value.value;

    // Get the Top border enabled
    const topBorder = card.enableTopBorder.value;

    // Get alignment
    const alignment = card.alignment.value.value;

    // Get the Bottom border enabled
    const bottomBorder = card.enableBottomBorder.value;

    // Grab the rowHeader and assign type
    const rowHeader = totalRow.children[0] as HTMLElement;

    // RowHeader Width
    const rowHeaderWidth = rowHeader.offsetWidth;

    // Apply a background image to the total row combining the rowHeaderBackground and the backgroundColor and rowHeader width
    totalRow.style.backgroundImage = `linear-gradient(to right, ${hex2rgba(
      rowHeaderBackground,
      card.headerOpacity.value
    )} ${rowHeaderWidth}px, ${hex2rgba(
      backgroundColor,
      card.opacity.value
    )} ${rowHeaderWidth}px)`;

    // Check if rowHeader includes invalid
    if (rowHeader.textContent.includes('Invalid')) {
      rowHeader.textContent = 'Total';
    }

    // Apply the custom label
    rowHeader.textContent =
      rowHeader.textContent === 'Total'
        ? totalTextContent
        : rowHeader.textContent;

    // Set the rowHeader values to the rowHeader
    rowHeader.style.fontFamily = rowHeaderFontFamily;
    rowHeader.style.fontSize = `${rowHeaderFontSize}px`;
    rowHeader.style.color = rowHeaderFontColor;
    rowHeader.style.justifyContent = rowHeaderAlignment;
    rowHeader.style.fontWeight = rowHeaderBold ? 'bold' : 'normal';
    rowHeader.style.fontStyle = rowHeaderItalic ? 'italic' : 'normal';
    rowHeader.style.textIndent = `${card.rowHeaderIndentation.value}px`;
    rowHeader.style.border = 'none';

    // Remove the border from the total row container numerous parents up
    totalRow.parentElement.parentElement.parentElement.style.border = 'none';

    // Set the top border
    totalRow.style.borderTop = topBorder
      ? `${borderWidth}px ${borderStyle} ${borderColor}`
      : 'none';

    // Set the bottom border
    totalRow.style.borderBottom = bottomBorder
      ? `${borderWidth}px ${borderStyle} ${borderColor}`
      : 'none';

    // Get the row node
    const rowNode = this.gridOptions.api.getPinnedBottomRow(0);

    // Set the height for multiple containers
    totalRow.parentElement.style.height = `${height}px`;
    totalRow.parentElement.parentElement.style.height = `${height}px`;
    totalRow.parentElement.parentElement.parentElement.style.height = `${height}px`;

    // Set via API to be sure
    rowNode.setRowHeight(height);

    // Loop through the children
    // Iterator to skip first child (As it is the rowHeader)
    let iterator = 0;
    // Loop through the category cells to find the correct row
    for (const child of Object(totalRow.children)) {
      if (iterator === 0) {
        iterator++;
        continue;
      }

      // Remove left and right borders of children
      child.style.border = 'none';

      // Child alignment
      child.style.justifyContent = alignment;

      // Set the row font
      child.style.fontFamily = font;

      // Set the row font size
      child.style.fontSize = `${fontSize}px`;

      // Set the row font color
      child.style.color = fontColor;

      // Set the bold
      child.style.fontWeight = bold ? 'bold' : 'normal';

      // Set the italic
      child.style.fontStyle = italic ? 'italic' : 'normal';

      // Set the indentation
      child.style.textIndent = `${card.indentation.value}px`;
    }
  }

  /**
   * Reorganies the grid after height changes via the API
   */
  private static onFinishHeightChange() {
    this.gridOptions.api.onRowHeightChanged();
  }
}

/**
 * HEX to RGBA converter
 */
function hex2rgba(hex, alpha = 1) {
  const [r, g, b] = hex.match(/\w\w/g).map((x) => parseInt(x, 16));
  return `rgba(${r},${g},${b},${alpha / 100})`;
}
