import React, { useState, useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import './App.css';

interface ExcelData {
  columns: string[];
  data: any[][];
}

interface ColumnMapping {
  templateColumn: string;
  dataColumn: string | null;
  defaultValue: string;
}

interface DOBMapping {
  month: string | null;
  day: string | null;
  year: string | null;
  complete: string | null;
}

const App: React.FC = () => {
  const [templateData, setTemplateData] = useState<ExcelData | null>(null);
  const [dataFile, setDataFile] = useState<ExcelData | null>(null);
  const [columnMappings, setColumnMappings] = useState<ColumnMapping[]>([]);
  const [dobMapping, setDobMapping] = useState<DOBMapping>({
    month: null,
    day: null,
    year: null,
    complete: null,
  });
  const [defaultValues, setDefaultValues] = useState<Record<string, string>>({});
  const [dragOverTarget, setDragOverTarget] = useState<string | null>(null);
  const [showInstructions, setShowInstructions] = useState(false);

  const parseExcelFile = (file: File): Promise<ExcelData> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
          
          const columns = jsonData[0] as string[];
          const dataRows = jsonData.slice(1) as any[][];
          
          resolve({ columns, data: dataRows });
        } catch (error) {
          reject(error);
        }
      };
      reader.readAsArrayBuffer(file);
    });
  };

  const onTemplateDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      parseExcelFile(acceptedFiles[0]).then((data) => {
        setTemplateData(data);
        // Initialize column mappings
        const mappings: ColumnMapping[] = data.columns.map(col => ({
          templateColumn: col,
          dataColumn: null,
          defaultValue: ''
        }));
        setColumnMappings(mappings);
      });
    }
  }, []);

  const onDataDrop = useCallback((acceptedFiles: File[]) => {
    if (acceptedFiles.length > 0) {
      parseExcelFile(acceptedFiles[0]).then((data) => {
        setDataFile(data);
      });
    }
  }, []);

  const { getRootProps: getTemplateRootProps, getInputProps: getTemplateInputProps, isDragActive: isTemplateDragActive } = useDropzone({
    onDrop: onTemplateDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    }
  });

  const { getRootProps: getDataRootProps, getInputProps: getDataInputProps, isDragActive: isDataDragActive } = useDropzone({
    onDrop: onDataDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    }
  });

  const handleDragStart = (e: React.DragEvent, column: string) => {
    e.dataTransfer.effectAllowed = 'move';
    // Allow scrolling during drag
    e.dataTransfer.setData('text/plain', column);
    
    // Add dragging class to body for scroll support
    document.body.classList.add('dragging');
  };

  const handleDragOver = (e: React.DragEvent, target: string) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    setDragOverTarget(target);
  };

  const handleDragLeave = () => {
    setDragOverTarget(null);
  };

  const handleDragEnd = () => {
    setDragOverTarget(null);
    // Remove dragging class
    document.body.classList.remove('dragging');
  };

  const handleDrop = (e: React.DragEvent, templateColumnIndex: number) => {
    e.preventDefault();
    setDragOverTarget(null);
    
    // Get the dragged column from dataTransfer
    const draggedColumnName = e.dataTransfer.getData('text/plain');
    
    if (draggedColumnName) {
      const newMappings = [...columnMappings];
      newMappings[templateColumnIndex].dataColumn = draggedColumnName;
      setColumnMappings(newMappings);
      
      console.log(`Template column ${templateColumnIndex} mapped to:`, draggedColumnName);
    }
    
    // Remove dragging class
    document.body.classList.remove('dragging');
  };

  const handleDOBDrop = (e: React.DragEvent, dobType: 'month' | 'day' | 'year' | 'complete') => {
    e.preventDefault();
    setDragOverTarget(null);
    
    // Get the dragged column from dataTransfer
    const draggedColumnName = e.dataTransfer.getData('text/plain');
    
    if (draggedColumnName) {
      const newDobMapping = { ...dobMapping };
      
      if (dobType === 'complete') {
        // Clear separate mappings when using complete date
        newDobMapping.month = null;
        newDobMapping.day = null;
        newDobMapping.year = null;
        newDobMapping.complete = draggedColumnName;
      } else {
        // Clear complete mapping when using separate columns
        newDobMapping.complete = null;
        newDobMapping[dobType] = draggedColumnName;
      }
      
      setDobMapping(newDobMapping);
      
      console.log(`DOB ${dobType} mapped to:`, draggedColumnName);
      console.log('Updated DOB mapping:', newDobMapping);
    }
    
    // Remove dragging class
    document.body.classList.remove('dragging');
  };

  const updateDefaultValue = (templateColumn: string, value: string) => {
    setDefaultValues(prev => ({
      ...prev,
      [templateColumn]: value
    }));
  };

  const ensureExcelCompatibility = (value: any, columnName: string): any => {
    // Handle different data types for Excel compatibility
    if (value === null || value === undefined || value === '') {
      return '';
    }
    
    // Convert to string first to handle all cases
    const stringValue = String(value).trim();
    
    // Handle numbers - ensure proper number format
    if (!isNaN(Number(stringValue)) && stringValue !== '') {
      const numValue = Number(stringValue);
      // Return as number for proper Excel formatting
      return numValue;
    }
    
    // Handle dates - format as proper Excel date
    if (columnName.toLowerCase().includes('date') || columnName.toLowerCase().includes('dob') || columnName.toLowerCase().includes('visit')) {
      const dateValue = new Date(stringValue);
      if (!isNaN(dateValue.getTime())) {
        // Return as properly formatted date string for Excel
        const year = dateValue.getFullYear();
        const month = String(dateValue.getMonth() + 1).padStart(2, '0');
        const day = String(dateValue.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
      }
    }
    
    // Handle boolean values
    if (stringValue.toLowerCase() === 'true' || stringValue.toLowerCase() === 'false') {
      return stringValue.toLowerCase() === 'true';
    }
    
    // Return as string for everything else - ensure proper encoding
    return stringValue;
  };

  const formatLastVisit = (dateValue: any): string => {
    if (!dateValue || dateValue === '' || dateValue === null || dateValue === undefined) {
      return '';
    }

    try {
      // Handle different date formats
      let date: Date;
      
      if (typeof dateValue === 'string') {
        // Try to parse various date formats
        if (dateValue.includes('/')) {
          // MM/DD/YYYY format
          date = new Date(dateValue);
        } else if (dateValue.includes('-')) {
          // YYYY-MM-DD format
          date = new Date(dateValue);
        } else {
          // Try parsing as is
          date = new Date(dateValue);
        }
      } else if (typeof dateValue === 'number') {
        // Excel date serial number
        date = new Date((dateValue - 25569) * 86400 * 1000);
      } else {
        date = new Date(dateValue);
      }

      if (isNaN(date.getTime())) {
        return '';
      }

      // Format as MM/DD/YYYY H:MM:SS AM/PM
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      const year = date.getFullYear();
      const hours = date.getHours();
      const minutes = String(date.getMinutes()).padStart(2, '0');
      const seconds = String(date.getSeconds()).padStart(2, '0');
      const ampm = hours >= 12 ? 'PM' : 'AM';
      const displayHours = hours % 12 || 12;

      return `${month}/${day}/${year} ${displayHours}:${minutes}:${seconds} ${ampm}`;
    } catch (error) {
      console.error('Error formatting date:', error);
      return '';
    }
  };

  const showExportInstructions = () => {
    setShowInstructions(true);
  };

  const confirmExport = () => {
    setShowInstructions(false);
    generateOutput();
  };

  const cancelExport = () => {
    setShowInstructions(false);
  };

  const generateOutput = async () => {
    if (!templateData || !dataFile) return;

    console.log('Default values:', defaultValues);
    console.log('DOB mapping:', dobMapping);
    console.log('Column mappings:', columnMappings);

    const maxRecords = 900;
    const totalRecords = dataFile.data.length;
    const totalFiles = Math.ceil(totalRecords / maxRecords);

    console.log(`Debug: totalRecords=${totalRecords}, maxRecords=${maxRecords}, totalFiles=${totalFiles}`);

    for (let fileIndex = 0; fileIndex < totalFiles; fileIndex++) {
      const startIndex = fileIndex * maxRecords;
      const endIndex = Math.min(startIndex + maxRecords, totalRecords);
      const results: any[][] = [];

      console.log(`Debug: Creating file ${fileIndex + 1}/${totalFiles}, records ${startIndex}-${endIndex-1}`);

      for (let i = startIndex; i < endIndex; i++) {
        const row: any[] = [];
        
        columnMappings.forEach((mapping, index) => {
          if (mapping.templateColumn === 'No') {
            // Auto-increment No column starting from 1 globally across all files
            row.push(i + 1);
          } else if (mapping.templateColumn === 'CustomerType') {
            // Default value for CustomerType
            row.push('Reward');
          } else if (mapping.templateColumn === 'Blacklist') {
            // Default value for Blacklist
            row.push(false);
          } else if (mapping.dataColumn) {
            const dataColumnIndex = dataFile.columns.indexOf(mapping.dataColumn);
            if (dataColumnIndex !== -1) {
              const cellValue = dataFile.data[i][dataColumnIndex];
              // Check if cell has meaningful data
              if (cellValue !== undefined && cellValue !== null && cellValue !== '' && cellValue !== ' ' && cellValue !== 'null' && cellValue !== 'NULL') {
                // Special formatting for LastVisit column
                if (mapping.templateColumn === 'LastVisit') {
                  const formattedDate = formatLastVisit(cellValue);
                  row.push(ensureExcelCompatibility(formattedDate, mapping.templateColumn));
                } else {
                  row.push(ensureExcelCompatibility(cellValue, mapping.templateColumn));
                }
              } else {
                // Use default value if data is missing or empty
                const defaultValue = defaultValues[mapping.templateColumn] || mapping.defaultValue || '';
                row.push(ensureExcelCompatibility(defaultValue, mapping.templateColumn));
              }
            } else {
              // Use default value if column not found
              const defaultValue = defaultValues[mapping.templateColumn] || mapping.defaultValue || '';
              row.push(ensureExcelCompatibility(defaultValue, mapping.templateColumn));
            }
          } else if (mapping.templateColumn === 'DateOfBirth' || mapping.templateColumn === 'Birthday') {
            // Handle DOB mapping - 2 options
            let dobValue = '';
            
            if (dobMapping.complete) {
              // Option 2: Complete date column
              const completeIndex = dataFile.columns.indexOf(dobMapping.complete);
              const completeDate = dataFile.data[i][completeIndex];
              
              if (completeDate !== undefined && completeDate !== null && completeDate !== '' && completeDate !== ' ') {
                // Convert mm-dd-yyyy to mm/dd/yyyy
                dobValue = completeDate.toString().replace(/-/g, '/');
              } else {
                dobValue = defaultValues['DateOfBirth'] ||defaultValues['Birthday']  || '';
              }
            } else if (dobMapping.month && dobMapping.day) {
              // Option 1: Separate columns
              const monthIndex = dataFile.columns.indexOf(dobMapping.month);
              const dayIndex = dataFile.columns.indexOf(dobMapping.day);
              
              const month = dataFile.data[i][monthIndex];
              const day = dataFile.data[i][dayIndex];
              
              // Check if month and day have meaningful data
              if (month !== undefined && month !== null && month !== '' && month !== ' ' && 
                  day !== undefined && day !== null && day !== '' && day !== ' ') {
                
                if (dobMapping.year) {
                  // Use provided year
                  const yearIndex = dataFile.columns.indexOf(dobMapping.year);
                  const year = dataFile.data[i][yearIndex];
                  if (year !== undefined && year !== null && year !== '' && year !== ' ') {
                    dobValue = `${month}/${day}/${year}`;
                  } else {
                    // Use current year if year data is missing
                    dobValue = `${month}/${day}/${new Date().getFullYear()}`;
                  }
                } else {
                  // Use current year if no year column is mapped
                  dobValue = `${month}/${day}/${new Date().getFullYear()}`;
                }
              } else {
                // Use default value if month or day is missing
                dobValue = defaultValues['DateOfBirth'] || defaultValues['Birthday']  || '';
              }
            } else {
              // Use default value if DOB mapping is incomplete
              dobValue = defaultValues['DateOfBirth'] || defaultValues['Birthday']  || '';
            }
            
            row.push(ensureExcelCompatibility(dobValue, mapping.templateColumn));
          } else {
            // Use default value for unmapped columns
            const defaultValue = defaultValues[mapping.templateColumn] || mapping.defaultValue || '';
            row.push(ensureExcelCompatibility(defaultValue, mapping.templateColumn));
          }
        });
        
        results.push(row);
      }

      // Create Excel file with minimal structure for maximum compatibility
      const ws = XLSX.utils.aoa_to_sheet(results);
      
      // Create workbook - minimal setup
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
      
      // Write with absolute minimal options
      const excelBuffer = XLSX.write(wb, { 
        bookType: 'xlsx', 
        type: 'array'
      });
      
      // Create blob with standard MIME type
      const blob = new Blob([excelBuffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      
      const fileName = `output_part_${String(fileIndex + 1).padStart(3, '0')}.xlsx`;
      console.log(`Debug: Saving file ${fileName} with ${results.length} records`);
      saveAs(blob, fileName);
      
      // Add delay between downloads to avoid browser limits
      if (fileIndex < totalFiles - 1) {
        await new Promise(resolve => setTimeout(resolve, 100));
      }
    }

    console.log(`Exported ${totalFiles} files with ${totalRecords} total records`);
  };

  return (
    <div className="app">
      <header className="app-header">
        <h1>Excel Data Filler</h1>
        <p>Upload template and data files, then drag columns to map them</p>
      </header>

      <div className="upload-section">
        <div className="upload-area">
          <h3>Template File</h3>
          <div {...getTemplateRootProps()} className={`dropzone ${isTemplateDragActive ? 'active' : ''}`}>
            <input {...getTemplateInputProps()} />
            <p>Drag & drop template Excel file here, or click to select</p>
          </div>
          {templateData && (
            <div className="file-info">
              <p>✅ Template loaded: {templateData.columns.length} columns</p>
            </div>
          )}
        </div>

        <div className="upload-area">
          <h3>Data File</h3>
          <div {...getDataRootProps()} className={`dropzone ${isDataDragActive ? 'active' : ''}`}>
            <input {...getDataInputProps()} />
            <p>Drag & drop data Excel file here, or click to select</p>
          </div>
          {dataFile && (
            <div className="file-info">
              <p>✅ Data loaded: {dataFile.data.length} records, {dataFile.columns.length} columns</p>
            </div>
          )}
        </div>
      </div>

      {templateData && dataFile && (
        <div className="mapping-section">
          <h3>Column Mapping</h3>
          <div className="mapping-container">
            <div className="data-columns">
              <h4>Data Columns (Drag these)</h4>
              <div className="column-list">
                {dataFile.columns.map((column, index) => (
                  <div
                    key={column}
                    className="column-item"
                    draggable
                    onDragStart={(e) => handleDragStart(e, column)}
                    onDragEnd={handleDragEnd}
                  >
                    {column}
                  </div>
                ))}
              </div>
            </div>

            <div className="template-columns">
              <h4>Template Columns</h4>
              
              {/* Template Titles Row */}
              {/* <div className="template-titles">
                {templateData.columns.map((column, index) => (
                  <div key={column} className="template-title">
                    <span className="title-text">{column}</span>
                    {column === 'No' ? (
                      <span className="auto-column">Auto-increment</span>
                    ) : columnMappings[index].dataColumn ? (
                      <span className="mapped-column">→ {columnMappings[index].dataColumn}</span>
                    ) : null}
                  </div>
                ))}
              </div> */}
              
              {/* Mapping Areas - Exclude No column */}
              <div className="template-mappings">
                {templateData.columns.filter(column => column !== 'No').map((column, index) => {
                  const originalIndex = templateData.columns.indexOf(column);
                  return (
                    <div key={column} className="template-mapping">
                      <div className="mapping-header">
                        <span>{column}</span>
                        {columnMappings[originalIndex].dataColumn ? (
                          <span className="mapped-column">→ {columnMappings[originalIndex].dataColumn}</span>
                        ) : null}
                      </div>
                  
                      {column === 'DateOfBirth' || column === 'Birthday' ? (
                        <div className="dob-mapping">
                          <div className="dob-options">
                            <div className="dob-option">
                              <h5>Option 1: Separate Columns</h5>
                              <div className="dob-slots">
                                <div 
                                  className={`dob-slot required ${dragOverTarget === 'dob-month' ? 'drag-over' : ''} ${dobMapping.month ? 'mapped' : ''}`}
                                  onDragOver={(e) => handleDragOver(e, 'dob-month')}
                                  onDragLeave={handleDragLeave}
                                  onDrop={(e) => handleDOBDrop(e, 'month')}
                                >
                                  <span>Month: {dobMapping.month || 'Drop here'} *</span>
                                </div>
                                <div 
                                  className={`dob-slot required ${dragOverTarget === 'dob-day' ? 'drag-over' : ''} ${dobMapping.day ? 'mapped' : ''}`}
                                  onDragOver={(e) => handleDragOver(e, 'dob-day')}
                                  onDragLeave={handleDragLeave}
                                  onDrop={(e) => handleDOBDrop(e, 'day')}
                                >
                                  <span>Day: {dobMapping.day || 'Drop here'} *</span>
                                </div>
                                <div 
                                  className={`dob-slot optional ${dragOverTarget === 'dob-year' ? 'drag-over' : ''} ${dobMapping.year ? 'mapped' : ''}`}
                                  onDragOver={(e) => handleDragOver(e, 'dob-year')}
                                  onDragLeave={handleDragLeave}
                                  onDrop={(e) => handleDOBDrop(e, 'year')}
                                >
                                  <span>Year: {dobMapping.year || 'Drop here (optional)'}</span>
                                </div>
                              </div>
                            </div>
                            <div className="dob-option">
                              <h5>Option 2: Complete Date Column</h5>
                              <div className="dob-complete-slot">
                                <div 
                                  className={`dob-slot complete ${dragOverTarget === 'dob-complete' ? 'drag-over' : ''} ${dobMapping.complete ? 'mapped' : ''}`}
                                  onDragOver={(e) => handleDragOver(e, 'dob-complete')}
                                  onDragLeave={handleDragLeave}
                                  onDrop={(e) => handleDOBDrop(e, 'complete')}
                                >
                                  <span>Complete Date: {dobMapping.complete || 'Drop here (mm-dd-yyyy or mm/dd/yyyy)'}</span>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      ) : (
                        <div 
                          className={`drop-zone ${dragOverTarget === `template-${originalIndex}` ? 'drag-over' : ''}`}
                          onDragOver={(e) => handleDragOver(e, `template-${originalIndex}`)}
                          onDragLeave={handleDragLeave}
                          onDrop={(e) => handleDrop(e, originalIndex)}
                        >
                          {columnMappings[originalIndex].dataColumn ? (
                            <span className="mapped">Mapped to: {columnMappings[originalIndex].dataColumn}</span>
                          ) : (
                            <span className="empty">Drop column here</span>
                          )}
                        </div>
                      )}
                      
                      <div className="default-value">
                        <input
                          type="text"
                          placeholder="Default value"
                          value={defaultValues[column] || ''}
                          onChange={(e) => updateDefaultValue(column, e.target.value)}
                        />
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      )}

      {templateData && dataFile && (
        <div className="export-section">
          <div className="debug-info">
            <h4>Debug Info:</h4>
            <p><strong>Default Values:</strong> {JSON.stringify(defaultValues)}</p>
            <p><strong>Auto Defaults:</strong> CustomerType="Reward", Blacklist=false</p>
            <p><strong>DOB Mapping:</strong> Month: {dobMapping.month || 'None'}, Day: {dobMapping.day || 'None'}, Year: {dobMapping.year || 'None'}, Complete: {dobMapping.complete || 'None'}</p>
            <p><strong>Mapped Columns:</strong> {columnMappings.filter(m => m.dataColumn).length} of {columnMappings.length}</p>
            <p><strong>Total Records:</strong> {dataFile.data.length} (will create {Math.ceil(dataFile.data.length / 900)} files)</p>
          </div>
          <button onClick={showExportInstructions} className="export-button">
            Export Multiple Excel Files (900 records per file, no headers)
          </button>
        </div>
      )}

      {/* Instructions Popup */}
      {showInstructions && (
        <div className="modal-overlay">
          <div className="modal-content">
            <h3>Xác nhận trên 18 tuổi</h3>
            <div className="instructions">
              <p><strong>Đọc trước hướng dẫn sử dụng trước khi dùng:</strong></p>
              <ol>
                <li>Khi tải xuống phải mở các file lên và bấm <strong>Ctrl + S</strong> xong là có thể dùng để import</li>
                <li>Nếu sau khi đã bấm <strong>Ctrl + S</strong> xong mà vẫn không thể import được thì tìm <strong>Brian Nguyen đẹp trai vô địch khắp vũ trụ từ team Tech Support</strong> để fix cho :D</li>
              </ol>
            </div>
            <div className="modal-buttons">
              <button onClick={cancelExport} className="cancel-button">
                Hủy
              </button>
              <button onClick={confirmExport} className="confirm-button">
                Đã Hiểu
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default App;