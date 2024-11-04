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

  private static processRowData(matrix, columnDefs, rowData) {
    let valueSourcesIndexBackUp = 0;

    matrix.rows.root.children.forEach((row) => {
      const rowObj = {};
      const identityHolder = [];

      if (row.isSubtotal !== true) {
        identityHolder.push(row);
      }
      rowObj['identity'] = identityHolder;

      // Case 1: Handle rows with values and no childIdentityFields
      if (
        Object.prototype.hasOwnProperty.call(row, 'values') &&
        !Object.prototype.hasOwnProperty.call(matrix.columns.root, 'childIdentityFields') &&
        !Object.prototype.hasOwnProperty.call(matrix.rows.root, 'childIdentityFields')
      ) {
        const rowValues = row.values as { [key: string]: any };
        let index = 0;

        Object.values(rowValues).forEach((value) => {
          const typedValue = value as { value: any; valueSourceIndex?: number };

          const valueFormatted = this.valueFormatterMatrix(
            typedValue.value,
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

      // Case 2: Handle rows with multiple values or children
      else if (Object.prototype.hasOwnProperty.call(row, 'values') || row.children) {
        const rowValues = row.values as { [key: string]: any };
        let index = 0;

        // Insert row header
        rowObj[Object(columnDefs)[index]['field']] = row.value;

        // Handle row expansion
        if (Object.prototype.hasOwnProperty.call(row, 'isCollapsed') || Object.prototype.hasOwnProperty.call(row, 'identity')) {
          this.nodesToExpand.push(row);
        }

        index++;

        // Handle row values or children
        try {
          Object.values(rowValues).forEach((value) => {
            const typedValue = value as { value: any; valueSourceIndex?: number };

            const valueFormatted = this.valueFormatterMatrix(
              typedValue.value,
              typedValue.valueSourceIndex === undefined
                ? valueSourcesIndexBackUp
                : typedValue.valueSourceIndex,
              rowValues
            );

            rowObj[Object(columnDefs)[index]['field']] = valueFormatted;
            index++;
            valueSourcesIndexBackUp++;
          });
        } catch {
          let lastItem = null;

          if (this.formattingSettings.expansionCard.expandUp.value === true) {
            lastItem = this.nodesToExpand.pop();
          }

          row.children.forEach((child) => {
            this.tempRowChildren.length = 0;
            this.traverseChildNodes(child, columnDefs, row.value, identityHolder);
            index++;
          });

          if (lastItem !== null) {
            this.nodesToExpand.push(lastItem);
          }
        }
      }

      // Case 3: Handle singular values with childIdentityFields
      else if (
        Object.prototype.hasOwnProperty.call(matrix.columns.root, 'childIdentityFields') ||
        Object.prototype.hasOwnProperty.call(matrix.rows.root, 'childIdentityFields')
      ) {
        rowObj[Object(columnDefs)[0]['field']] = row.value;
      }

      // Push the processed row object to rowData
      rowData.push(rowObj);
    });
  }

  private static finalizeRowData(rowData, dynamicHeader, matrix) {
    // Handle row data and expansion upwards
    if (!this.expandUpwards && this.rowChildrenNodes.length > 0) {
      rowData.length = 0;
      const tempHolderNodes = [...this.rowChildrenNodes];
      tempHolderNodes.unshift(tempHolderNodes.pop());
      this.rowChildrenNodesSorted = [...this.rowChildrenNodesSorted, ...tempHolderNodes];
      this.rowChildrenNodes = [];

      this.rowChildrenNodesSorted.forEach((node) => rowData.push(node));
    } else if (this.rowChildrenNodes.length > 0) {
      rowData.length = 0;
      this.rowChildrenNodes.forEach((child) => rowData.push(child));
    }

    if (rowData.length !== this.rowChildrenNodesSorted.length && this.rowChildrenNodesSorted.length > 0) {
      rowData.pop();
    }

    // Handle the "Total" row logic
    const lengthOfRowData = rowData.length;

    if (rowData[lengthOfRowData - 1][dynamicHeader as keyof typeof rowData] === undefined) {
      if (
        Object.prototype.hasOwnProperty.call(matrix.columns.root, 'childIdentityFields') ||
        Object.prototype.hasOwnProperty.call(matrix.rows.root, 'childIdentityFields')
      ) {
        rowData[lengthOfRowData - 1][dynamicHeader as keyof typeof rowData] = 'Total';
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

    //let colId = 0;
    const colId = 0;

    this.handleSingularMeasureMatrix(matrix, dataView, columnDefs, colId, dynamicHeader, gridCellStyling, categoryCellStyling, rowHeadersCard);

    this.defineColumns(matrix, dataView, columnDefs, colId, dynamicHeader, gridCellStyling, categoryCellStyling, rowHeadersCard);

    const rowData: object[] = [];
    this.processRowData(matrix, columnDefs, rowData);
    this.finalizeRowData(rowData, dynamicHeader, matrix);
    //this.buildRowData(matrix, columnDefs, rowData, dynamicHeader);

    //const total = rowData.pop();
    return this.finalizeGrid(columnDefs, rowData, dynamicHeader);
  }

  private static processChildRow(child, columnDefs, parentHeader, identityHolder) {
    // Identity holder in a separate array to avoid reference duplication
    const newIdentityHolder = [...identityHolder];

    // Create the rowObj
    const rowObj = {};

    // Deconstruct row into values
    const rowValues = child.values;

    // Index iterator
    let index = 0;

    // Check if child.value (aka the header) is undefined, if so grab the parentHeader
    //let rowHeader = child.value !== undefined ? child.value : parentHeader;
    const rowHeader = child.value !== undefined ? child.value : parentHeader;

    // Insert row header
    rowObj[Object(columnDefs)[index]['field']] = String(rowHeader);

    // Add the nodes to the nodesToExpand array so expand buttons can be programmatically added & for expansion identification
    if (
      Object.prototype.hasOwnProperty.call(child, 'isCollapsed') ||
      Object.prototype.hasOwnProperty.call(child, 'identity')
    ) {
      this.nodesToExpand.push(child);
    }

    // If identity holder is not subTotal, push it
    if (child.isSubtotal !== true) {
      newIdentityHolder.push(child);
    }

    // Push the identity holder into the rowObj
    rowObj['identity'] = newIdentityHolder;

    // Increment index
    index++;

    // Valuesource index back up (For calculation groups)
    let valueSourcesIndexBackUp = 0;

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
    } catch {
      // If there's an error, it means children need to be processed recursively
      rowObj['hasChildren'] = true;
    }

    return { rowObj, rowHeader, newIdentityHolder };
  }

  private static async traverseAndExpandChildren(child, columnDefs, rowHeader, identityHolder) {
    const { rowObj, newIdentityHolder } = this.processChildRow(child, columnDefs, rowHeader, identityHolder);

    if (!rowObj['hasChildren']) {
      if (this.expandUpwards === false) {
        // Manage tempRowChildren and rowChildrenNodes for expansion downwards
        this.tempRowChildren.push(rowObj);
        this.rowChildrenNodes.push(rowObj);

        const tempRowChildren = [...this.tempRowChildren];
        const lastItem = tempRowChildren.pop();
        tempRowChildren.unshift(lastItem);

        const lengthOfRowChildrenNodes = this.rowChildrenNodes.length;
        const lengthOfTempNodes = tempRowChildren.length;
        const difference = lengthOfRowChildrenNodes - lengthOfTempNodes;

        this.rowChildrenNodes.splice(difference, lengthOfTempNodes);
        this.rowChildrenNodes = [...this.rowChildrenNodes, ...tempRowChildren];
      } else {
        this.rowChildrenNodes.push(rowObj);
      }
    } else {
      // Recursive case: process children and handle upward expansion
      const tempRowChildrenContained = [...this.tempRowChildren];

      if (this.expandUpwards === false) {
        this.tempRowChildren.length = 0;
      }

      let lastItem = null;
      if (this.formattingSettings.expansionCard.expandUp.value === true) {
        lastItem = this.nodesToExpand.pop();
      }

      child.children.forEach((subChild) => {
        this.traverseAndExpandChildren(subChild, columnDefs, rowHeader, newIdentityHolder);
      });

      if (lastItem !== null) {
        this.nodesToExpand.push(lastItem);
      }

      if (this.expandUpwards === false) {
        tempRowChildrenContained.reverse();
        this.tempRowChildren.unshift(this.tempRowChildren.pop());
        this.tempRowChildren = [
          ...tempRowChildrenContained,
          ...this.tempRowChildren,
        ];
      }
    }
  }

  /**
 * Recursive function to traverse child nodes
 */
  private static async traverseChildNodes(child, columnDefs, parentHeader, identityHolder) {
    const { rowObj, rowHeader, newIdentityHolder } = this.processChildRow(child, columnDefs, parentHeader, identityHolder);

    this.traverseAndExpandChildren(child, columnDefs, rowHeader, newIdentityHolder);
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
   * Checks if expansion requirements are met before formatting rows.
   */
  private static checkExpansionRequirements(): boolean {
    if (
      !Object.prototype.hasOwnProperty.call(this.dataView.matrix.columns.root, 'childIdentityFields') &&
      !Object.prototype.hasOwnProperty.call(this.dataView.matrix.rows.root, 'childIdentityFields')
    ) {
      return false;
    }

    try {
      if (
        !Object.prototype.hasOwnProperty.call(this.dataView.matrix.rows.levels[0], 'canBeExpanded')
      ) {
        return false;
      }
    } catch {
      return false;
    }

    return true;
  }

  /**
 * Applies formatting to each expanded row based on expansion settings.
 */
  private static applyRowFormatting() {
    const expansionCard = this.formattingSettings.expansionCard;
    const rows = document.querySelectorAll('.ag-row');

    for (const rowHeader of Object(rows)) {
      const rowId = rowHeader.getAttribute('row-id');
      const rowNodeData = this.nodesToExpand[rowId];
      const rowHeaderTextcontent = rowHeader.children[0].textContent;
      const rowHeaderDIV = rowHeader.children[0];

      let indentation = expansionCard.indentationValue.value;

      try {
        if (rowHeaderTextcontent === 'Total') {
          if (expansionCard.enableIndentation.value && expansionCard.enableButtons.value) {
            rowHeaderDIV.style.textIndent = `${indentation}px`;
            rowHeaderDIV.style.textAlign = 'start';
          }
          continue;
        }

        let nodeLevel = rowNodeData.level === 0 ? 1 : rowNodeData.level + 1;
        if (!expansionCard.enableButtons.value) nodeLevel -= 1;

        indentation = parseInt(nodeLevel) * indentation;

        if (expansionCard.enableIndentation.value) {
          rowHeaderDIV.style.textIndent = `${indentation}px`;
        }

        if (!rowNodeData.isCollapsed) {
          rowHeader.classList.add('expandedRow');

          for (const child of rowHeader.children) {
            child.style.fontWeight = expansionCard.enableBold.value ? 'bold' : 'normal';
            child.style.fontStyle = expansionCard.enableItalic.value ? 'italic' : 'normal';
            child.style.justifyContent = expansionCard.alignment.value.value;
            child.style.fontFamily = expansionCard.fontFamily.value;
            child.style.fontSize = `${expansionCard.fontSize.value}px`;
            child.style.color = expansionCard.fontColor.value.value;
          }

          const borderWidth = expansionCard.borderWidth.value;
          const borderColor = hex2rgba(
            expansionCard.borderColor.value.value,
            expansionCard.borderOpacity.value
          );
          const borderStyle = expansionCard.borderStyle.value.value;

          rowHeader.style.borderTop = expansionCard.enableTopBorder.value
            ? `${borderWidth}px ${borderStyle} ${borderColor}`
            : 'none';
          rowHeader.style.borderBottom = expansionCard.enableBottomBorder.value
            ? `${borderWidth}px ${borderStyle} ${borderColor}`
            : 'none';

          this.gridOptions.api.getRowNode(rowId).setRowHeight(expansionCard.height.value);
          rowHeader.style.backgroundImage = `linear-gradient(to right, ${hex2rgba(
            expansionCard.backgroundColor.value.value,
            expansionCard.opacity.value
          )} 100%, white 100%)`;

          const button = rowHeaderDIV.querySelector('.expandButton');
          if (button) button.textContent = '-';
        }
      } catch (e) {
        console.error('Error applying row formatting:', e);
      }
    }
  }

  /**
 * This function formats the expanded rows.
 */
  private static formatExpandedRows() {
    // Run initial checks to prevent unnecessary formatting
    if (!this.checkExpansionRequirements()) return;

    // Apply row formatting
    this.applyRowFormatting();
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
 * Extracts formatting settings from the expansion card.
 */
  private static extractFormattingSettings(card) {
    return {
      applicableRows: card.savedName.value.split(','),
      rowHeaderFontFamily: card.rowHeaderFontFamily.value,
      height: card.height.value,
      rowHeaderFontColor: card.rowHeaderFontColor.value.value,
      rowHeaderFontSize: card.rowHeaderFontSize.value,
      rowHeaderBold: card.rowHeaderBold.value,
      rowHeaderItalic: card.rowHeaderItalic.value,
      rowHeaderAlignment: card.rowHeaderAlignment.value.value,
      rowHeaderBackground: card.rowHeaderBackground.value.value,
      font: card.fontFamily.value,
      fontSize: card.fontSize.value,
      fontColor: card.fontColor.value.value,
      bold: card.enableBold.value,
      italic: card.enableItalic.value,
      backgroundColor: card.backgroundColor.value.value,
      borderWidth: card.borderWidth.value,
      borderColor: hex2rgba(card.borderColor.value.value, card.borderOpacity.value),
      borderStyle: card.borderStyle.value.value,
      topBorder: card.enableTopBorder.value,
      bottomBorder: card.enableBottomBorder.value,
      alignment: card.alignment.value.value,
      rowHeaderIndentation: card.rowHeaderIndentation.value,
      indentation: card.indentation.value,
      headerOpacity: card.headerOpacity.value,
      opacity: card.opacity.value,
    };
  }

  /**
 * Applies formatting to a specific row.
 */
  private static applyFormattingToRow(row, rowHeader, rowSettings, rowHeaderWidth) {
    // Set row height if applicable
    const rowId = row.getAttribute('row-id');
    const rowNode = this.gridOptions.api.getRowNode(rowId);
    if (rowNode) rowNode.setRowHeight(rowSettings.height);

    // Set top and bottom border
    row.style.borderTop = rowSettings.topBorder
      ? `${rowSettings.borderWidth}px ${rowSettings.borderStyle} ${rowSettings.borderColor}`
      : 'none';

    row.style.borderBottom = rowSettings.bottomBorder
      ? `${rowSettings.borderWidth}px ${rowSettings.borderStyle} ${rowSettings.borderColor}`
      : 'none';

    // Apply child-specific styles
    for (const child of row.children) {
      child.style.justifyContent = rowSettings.alignment;
      child.style.fontFamily = rowSettings.font;
      child.style.fontSize = `${rowSettings.fontSize}px`;
      child.style.color = rowSettings.fontColor;
      child.style.fontWeight = rowSettings.bold ? 'bold' : 'normal';
      child.style.fontStyle = rowSettings.italic ? 'italic' : 'normal';
      child.style.textIndent = `${rowSettings.indentation}px`;
      child.style.backgroundColor = 'RGBA(0,0,0,0)';
    }

    // Set background gradient
    row.style.backgroundImage = `linear-gradient(to right, ${hex2rgba(
      rowSettings.rowHeaderBackground,
      rowSettings.headerOpacity
    )} ${rowHeaderWidth}px, ${hex2rgba(
      rowSettings.backgroundColor,
      rowSettings.opacity
    )} ${rowHeaderWidth}px)`;

    // Apply row header-specific styles
    rowHeader.style.fontFamily = rowSettings.rowHeaderFontFamily;
    rowHeader.style.fontSize = `${rowSettings.rowHeaderFontSize}px`;
    rowHeader.style.color = rowSettings.rowHeaderFontColor;
    rowHeader.style.justifyContent = rowSettings.rowHeaderAlignment;
    rowHeader.style.fontWeight = rowSettings.rowHeaderBold ? 'bold' : 'normal';
    rowHeader.style.fontStyle = rowSettings.rowHeaderItalic ? 'italic' : 'normal';
    rowHeader.style.textIndent = `${rowSettings.rowHeaderIndentation}px`;
  }

  /**
 * Format specific rows.
 */
  private static formatSpecificRows() {
    const categoryCells = document.querySelectorAll('.categoryCell');
    if (!categoryCells) return;

    for (const card of this.formattingSettings.cards) {
      if (!card.name.includes('specificRow') || !card.enableCard.value || !card.savedName.value) {
        continue;
      }

      const rowSettings = this.extractFormattingSettings(card);

      for (const displayName of rowSettings.applicableRows) {
        if (
          displayName === 'Total' ||
          displayName === this.formattingSettings.totalCard.savedName.value
        ) {
          continue;
        }

        for (const rowHeader of Object(categoryCells)) {
          const rowTextContent = rowHeader.textContent.replace('+', '').replace('-', '');

          if (rowTextContent !== displayName.replace('+', '').replace('-', '')) {
            continue;
          }

          const row = rowHeader.parentElement;
          const rowHeaderWidth = rowHeader.offsetWidth;

          this.applyFormattingToRow(row, rowHeader, rowSettings, rowHeaderWidth);
        }
      }
    }
  }

  /**
   * Extracts formatting settings for a specific column from a card.
   */
  private static extractColumnFormattingSettings(card) {
    return {
      applicableColumns: card.savedName.value.split(','),
      alignment: card.columnAlignment.value.value,
      backgroundColor: hex2rgba(card.columnBackgroundColor.value.value, card.valuesOpacity.value),
      bold: card.columnBold.value,
      italic: card.columnItalic.value,
      fontColor: card.columnFontColor.value.value,
      font: card.columnFontFamily.value,
      fontSize: card.columnFontSize.value,
      columnHeaderAlignment: card.columnHeaderAlignment.value.value,
      columnHeaderBackgroundColor: hex2rgba(card.columnHeaderBackgroundColor.value.value, card.opacity.value),
      columnHeaderBold: card.columnHeaderBold.value,
      columnHeaderFontColor: card.columnHeaderFontColor.value.value,
      columnHeaderFontFamily: card.columnHeaderFontFamily.value,
      columnHeaderFontSize: card.columnHeaderFontSize.value,
      columnHeaderItalic: card.columnHeaderItalic.value,
      columnWidth: card.columnWidth.value,
      enableLeftBorder: card.enableLeftBorder.value,
      enableRightBorder: card.enableRightBorder.value,
      borderColor: hex2rgba(card.borderColor.value.value, card.borderOpacity.value),
      borderStyle: card.borderStyle.value.value,
      borderWidth: card.borderWidth.value,
    };
  }

  /**
   * Applies formatting to a specific column and its header.
   */
  private static applyFormattingToColumn(header, colId, columnSettings) {
    //const columnChildren = document.querySelectorAll(`[col-id="${colId}"]`);
    const columnChildren = Object(document.querySelectorAll(`[col-id="${colId}"]`));

    // Apply styles to each cell in the column
    for (const child of columnChildren) {
      child.style.justifyContent = columnSettings.alignment;
      child.style.backgroundColor = columnSettings.backgroundColor;
      child.style.fontFamily = columnSettings.font;
      child.style.fontSize = `${columnSettings.fontSize}px`;
      child.style.color = columnSettings.fontColor;
      child.style.fontWeight = columnSettings.bold ? 'bold' : 'normal';
      child.style.fontStyle = columnSettings.italic ? 'italic' : 'normal';
      child.style.borderLeft = columnSettings.enableLeftBorder
        ? `${columnSettings.borderWidth}px ${columnSettings.borderStyle} ${columnSettings.borderColor}`
        : 'none';
      child.style.borderRight = columnSettings.enableRightBorder
        ? `${columnSettings.borderWidth}px ${columnSettings.borderStyle} ${columnSettings.borderColor}`
        : 'none';
      child.style.borderTop = 'none';
      child.style.borderBottom = 'none';
    }

    // Apply styles to the header
    header.style.fontFamily = columnSettings.columnHeaderFontFamily;
    header.style.fontStyle = columnSettings.columnHeaderItalic ? 'italic' : 'normal';
    header.style.fontWeight = columnSettings.columnHeaderBold ? 'bold' : 'normal';
    header.style.fontSize = `${columnSettings.columnHeaderFontSize}px`;
    header.style.color = columnSettings.columnHeaderFontColor;

    const headerParent = header.parentElement.parentElement.parentElement.parentElement;
    headerParent.style.backgroundColor = columnSettings.columnHeaderBackgroundColor;
    headerParent.style.display = 'flex';
    header.style.justifyContent = columnSettings.columnHeaderAlignment;

    // Apply the column width via gridOptions
    this.gridOptions.columnApi.setColumnWidths([{ key: colId, newWidth: columnSettings.columnWidth }]);
  }

  /**
   * Formats specific columns based on settings.
   */
  private static formatSpecificColumns() {
    //const headerCells = document.querySelectorAll('.ag-header-cell-text');
    const headerCells = Object(document.querySelectorAll('.ag-header-cell-text'));


    for (const card of this.formattingSettings.cards) {
      if (!card.name.includes('specificColumn') || !card.enableCard.value) {
        continue;
      }

      const columnSettings = this.extractColumnFormattingSettings(card);

      for (const displayName of columnSettings.applicableColumns) {
        for (const header of headerCells) {
          if (header.textContent !== displayName) {
            continue;
          }

          const headerParent = header.parentElement.parentElement.parentElement.parentElement;
          const colId = headerParent.getAttribute('col-id');

          this.applyFormattingToColumn(header, colId, columnSettings);
        }
      }
    }
  }


  /**
   * Extracts formatting settings for the Total row from the card.
   */
  private static extractTotalRowSettings(card) {
    return {
      totalTextContent: card.savedName?.value || 'Total',
      rowHeaderFontFamily: card.rowHeaderFontFamily.value,
      height: card.height.value,
      rowHeaderFontColor: card.rowHeaderFontColor.value.value,
      rowHeaderFontSize: card.rowHeaderFontSize.value,
      rowHeaderBold: card.rowHeaderBold.value,
      rowHeaderItalic: card.rowHeaderItalic.value,
      rowHeaderAlignment: card.rowHeaderAlignment.value.value,
      rowHeaderBackground: card.rowHeaderBackground.value.value,
      font: card.fontFamily.value,
      fontSize: card.fontSize.value,
      fontColor: card.fontColor.value.value,
      bold: card.enableBold.value,
      italic: card.enableItalic.value,
      backgroundColor: card.backgroundColor.value.value,
      borderWidth: card.borderWidth.value,
      borderColor: hex2rgba(card.borderColor.value.value, card.borderOpacity.value),
      borderStyle: card.borderStyle.value.value,
      topBorder: card.enableTopBorder.value,
      bottomBorder: card.enableBottomBorder.value,
      alignment: card.alignment.value.value,
      rowHeaderIndentation: card.rowHeaderIndentation.value,
      headerOpacity: card.headerOpacity.value,
      opacity: card.opacity.value,
      indentation: card.indentation.value,
    };
  }

  /**
 * Applies formatting to the Total row.
 */
  private static applyFormattingToTotalRow(totalRow, rowHeader, rowSettings) {
    const rowHeaderWidth = rowHeader.offsetWidth;

    // Apply background gradient to the total row
    totalRow.style.backgroundImage = `linear-gradient(to right, ${hex2rgba(
      rowSettings.rowHeaderBackground,
      rowSettings.headerOpacity
    )} ${rowHeaderWidth}px, ${hex2rgba(
      rowSettings.backgroundColor,
      rowSettings.opacity
    )} ${rowHeaderWidth}px)`;

    // Apply custom label
    if (rowHeader.textContent.includes('Invalid')) rowHeader.textContent = 'Total';
    rowHeader.textContent = rowHeader.textContent === 'Total'
      ? rowSettings.totalTextContent
      : rowHeader.textContent;

    // Apply rowHeader-specific styles
    rowHeader.style.fontFamily = rowSettings.rowHeaderFontFamily;
    rowHeader.style.fontSize = `${rowSettings.rowHeaderFontSize}px`;
    rowHeader.style.color = rowSettings.rowHeaderFontColor;
    rowHeader.style.justifyContent = rowSettings.rowHeaderAlignment;
    rowHeader.style.fontWeight = rowSettings.rowHeaderBold ? 'bold' : 'normal';
    rowHeader.style.fontStyle = rowSettings.rowHeaderItalic ? 'italic' : 'normal';
    rowHeader.style.textIndent = `${rowSettings.rowHeaderIndentation}px`;
    rowHeader.style.border = 'none';

    // Set the top and bottom border
    totalRow.style.borderTop = rowSettings.topBorder
      ? `${rowSettings.borderWidth}px ${rowSettings.borderStyle} ${rowSettings.borderColor}`
      : 'none';
    totalRow.style.borderBottom = rowSettings.bottomBorder
      ? `${rowSettings.borderWidth}px ${rowSettings.borderStyle} ${rowSettings.borderColor}`
      : 'none';

    // Set row height
    const rowNode = this.gridOptions.api.getPinnedBottomRow(0);
    rowNode.setRowHeight(rowSettings.height);
    totalRow.parentElement.style.height = `${rowSettings.height}px`;
    totalRow.parentElement.parentElement.style.height = `${rowSettings.height}px`;
    totalRow.parentElement.parentElement.parentElement.style.height = `${rowSettings.height}px`;

    // Apply cell-specific styles to each child in the Total row
    let iterator = 0;
    for (const child of totalRow.children) {
      if (iterator++ === 0) continue;

      child.style.border = 'none';
      child.style.justifyContent = rowSettings.alignment;
      child.style.fontFamily = rowSettings.font;
      child.style.fontSize = `${rowSettings.fontSize}px`;
      child.style.color = rowSettings.fontColor;
      child.style.fontWeight = rowSettings.bold ? 'bold' : 'normal';
      child.style.fontStyle = rowSettings.italic ? 'italic' : 'normal';
      child.style.textIndent = `${rowSettings.indentation}px`;
    }
  }

  /**
   * Formats the Total row.
   */
  private static formatTotal() {
    const totalRow = document.querySelector('.ag-row-pinned') as HTMLElement;
    if (!totalRow) return;

    const card = this.formattingSettings.totalCard;
    if (!card.enableCard.value) {
      totalRow.parentElement.parentElement.parentElement.style.display = 'none';
      return;
    }

    const rowSettings = this.extractTotalRowSettings(card);
    const rowHeader = totalRow.children[0] as HTMLElement;

    this.applyFormattingToTotalRow(totalRow, rowHeader, rowSettings);
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
