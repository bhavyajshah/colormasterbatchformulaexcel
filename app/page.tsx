'use client';
import { useState, useMemo, useEffect } from 'react';
import { useQuery } from '@tanstack/react-query';
import * as XLSX from 'xlsx';
import { ChevronLeft, ChevronRight, FileSpreadsheet, Search } from 'lucide-react';
import Auth from './auth';

type CellValue = string | number | boolean | null | undefined;
type ExcelRow = CellValue[];

interface ExcelData {
  [sheetName: string]: ExcelRow[];
}

const ITEMS_PER_PAGE = 10;

const fetchExcelData = async (): Promise<ExcelData> => {
  const response = await fetch('/Merge Source Data of BOM Header & BOM Item.xlsx');
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  
  const data: ExcelData = {};
  workbook.SheetNames.forEach((sheetName) => {
    const worksheet = workbook.Sheets[sheetName];
    // Use defval to fill empty cells and ensure we get all columns
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
      header: 1, 
      defval: null,
      raw: false 
    });
    
    // For BOM Header: skip first 4 rows (0,1,2,3) and use row 4 as header to remove ETE;80;0;C;80;0 row
    // For BOM Item: skip first 1 row (0) and use row 1 as header, start data from row 2
    let headerRowIndex = 3;
    let dataStartIndex = 4;
    
    if (sheetName.toLowerCase().includes('item')) {
      headerRowIndex = 0;
      dataStartIndex = 1;
    }
    
    const headers = jsonData[headerRowIndex] as ExcelRow || [];
    const dataRows = jsonData.slice(dataStartIndex) as ExcelRow[];
    
    // Select specific columns based on sheet type
    let selectedHeaders: ExcelRow;
    let selectedDataRows: ExcelRow[];
    
    if (sheetName.toLowerCase().includes('item')) {
      // BOM Item: Show columns A, B, D, H, M, N, AP (indices 0, 1, 3, 7, 12, 13, 41)
      const columnIndices = [0, 1, 3, 7, 12, 13, 41];
      selectedHeaders = columnIndices.map(i => headers[i] || '');
      selectedDataRows = dataRows.map(row => 
        columnIndices.map(i => row[i] || null)
      );
    } else {
      // BOM Header: Show columns A, B, D, H, I, J (indices 0, 1, 3, 7, 8, 9)
      const columnIndices = [0, 1, 3, 7, 8, 9];
      selectedHeaders = columnIndices.map(i => headers[i] || '');
      selectedDataRows = dataRows.map(row => 
        columnIndices.map(i => row[i] || null)
      );
    }
    
    const processedData = [selectedHeaders, ...selectedDataRows];
    data[sheetName] = processedData;
  });
  
  return data;
};

export default function Home() {
  const [activeTab, setActiveTab] = useState<string>('');
  const [currentPage, setCurrentPage] = useState(1);
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedColor, setSelectedColor] = useState<string>('');
  const [calculationResults, setCalculationResults] = useState<ExcelRow[][]>([]);
  const [colorSearchTerm, setColorSearchTerm] = useState('');
  const [isColorDropdownOpen, setIsColorDropdownOpen] = useState(false);
  const [isCalculating, setIsCalculating] = useState(false);
  const [recentColors, setRecentColors] = useState<string[]>([]);
  const [favorites, setFavorites] = useState<string[]>([]);

  const { data: excelData, isLoading, error } = useQuery({
    queryKey: ['excelData'],
    queryFn: fetchExcelData
  });

  // Handle setting initial active tab when data loads
  useEffect(() => {
    if (excelData && !activeTab && Object.keys(excelData).length > 0) {
      setActiveTab(Object.keys(excelData)[0]);
    }
  }, [excelData, activeTab]);

  const currentData = useMemo(() => {
    if (!excelData || !activeTab || !excelData[activeTab]) return [];
    
    const data = excelData[activeTab];
    if (data.length === 0) return [];
    
    // Get all data without filtering empty rows - show everything including N/A
    const allData = data.slice(1);
    
    // Filter data based on search term only
    const filteredData = allData.filter((row: ExcelRow) =>
      searchTerm.trim() === '' || row.some((cell: CellValue) =>
        cell?.toString().toLowerCase().includes(searchTerm.toLowerCase())
      )
    );
    
    return filteredData;
  }, [excelData, activeTab, searchTerm]);

  const paginatedData = useMemo(() => {
    const startIndex = (currentPage - 1) * ITEMS_PER_PAGE;
    return currentData.slice(startIndex, startIndex + ITEMS_PER_PAGE);
  }, [currentData, currentPage]);

  const totalPages = Math.ceil(currentData.length / ITEMS_PER_PAGE);

  const headers = useMemo(() => {
    if (!excelData || !activeTab || !excelData[activeTab] || excelData[activeTab].length === 0) {
      return [];
    }
    return excelData[activeTab][0] || [];
  }, [excelData, activeTab]);

  // Get BOM Header colors for dropdown
  const bomHeaderColors = useMemo(() => {
    if (!excelData || !excelData['BOM Header']) return [];
    const headerData = excelData['BOM Header'];
    
    // Debug: log the header data structure
    console.log('BOM Header data:', headerData.slice(0, 5));
    
    // Get all rows including potential missing ones - check different column indices
    const allColors = headerData.slice(1).map((row, index) => {
      console.log(`Row ${index}:`, row);
      return {
        code: row[0] as string, // Column A - Code
        color: (row[5] || row[4] || row[3] || row[2] || row[1]) as string  // Try multiple columns for color name
      };
    }).filter(item => item.code && item.code.trim() !== '');
    
    console.log('Processed colors:', allColors);
    
    // Also check BOM Item for any codes not in header
    if (excelData['BOM Item']) {
      const itemData = excelData['BOM Item'];
      const itemCodes = new Set();
      itemData.slice(1).forEach(row => {
        const code = row[0] as string;
        if (code && code.trim() !== '') {
          itemCodes.add(code);
        }
      });
      
      // Add any missing codes from BOM Item
      itemCodes.forEach(code => {
        if (!allColors.find(item => item.code === code)) {
          allColors.push({
            code: code as string,
            color: `Color for ${code}` // Default name if not found in header
          });
        }
      });
    }
    
    return allColors;
  }, [excelData]);

  // Get filtered colors based on search term
  const filteredColors = useMemo(() => {
    return bomHeaderColors.filter(item => 
      item.code.toLowerCase().includes(colorSearchTerm.toLowerCase()) ||
      item.color.toLowerCase().includes(colorSearchTerm.toLowerCase())
    );
  }, [bomHeaderColors, colorSearchTerm]);

  // Export to CSV function - handle multiple formulations
  const exportToCSV = () => {
    if (calculationResults.length === 0) return;
    
    const headers = ['Code', 'Unit', 'Qty', 'Description', 'Weight (kg)', 'Unit', 'Type'];
    let csvContent = headers.join(',') + '\n';
    
    calculationResults.forEach((formulation, index) => {
      csvContent += `\n"=== Formulation ${index + 1} ===","","","","","",""\n`;
      formulation.forEach(row => {
        csvContent += row.map(cell => `"${String(cell || '')}"`).join(',') + '\n';
      });
    });
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    link.setAttribute('href', url);
    link.setAttribute('download', `formulations_${selectedColor}_all.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Pre-calculate all formulations for better performance - handle duplicates separately
  const allFormulations = useMemo(() => {
    if (!excelData || !excelData['BOM Item']) return {};
    
    const bomItemData = excelData['BOM Item'];
    const formulations: { [key: string]: ExcelRow[][] } = {}; // Array of arrays for multiple formulations
    
    bomHeaderColors.forEach(colorItem => {
      const colorRows = bomItemData.slice(1).filter(row => row[0] === colorItem.code);
      
      if (colorRows.length > 0) {
        // Group rows by sequential 7-ingredient formulations that total ~100kg
        const allCalculatedFormulations: ExcelRow[][] = [];
        
        // Process rows in groups of 7 (typical formulation size)
        for (let i = 0; i < colorRows.length; i += 7) {
          const formulationGroup = colorRows.slice(i, i + 7);
          
          if (formulationGroup.length > 0) {
            const totalWeight = formulationGroup.reduce((sum, row) => {
              const weight = parseFloat(row[4] as string) || 0;
              return sum + weight;
            }, 0);
            
            const calculatedRows = formulationGroup.map(row => {
              const originalWeight = parseFloat(row[4] as string) || 0;
              const proportionalWeight = totalWeight > 0 ? (originalWeight / totalWeight) * 100 : 0;
              
              return [
                row[0], // Code
                row[1], // UM03
                row[2], // 1
                row[3], // Color description
                proportionalWeight.toFixed(3), // Calculated weight for 100kg
                row[5], // KG
                row[6]  // RM02
              ];
            });
            
            allCalculatedFormulations.push(calculatedRows);
          }
        }
        
        formulations[colorItem.code] = allCalculatedFormulations;
      }
    });
    
    return formulations;
  }, [excelData, bomHeaderColors]);

  // Fast formulation lookup - handle multiple formulations as separate arrays
  const calculateFormulation = (colorCode: string) => {
    const formulations = allFormulations[colorCode];
    if (formulations && formulations.length > 0) {
      // Keep formulations separate for individual table rendering
      setCalculationResults(formulations);
    } else {
      setCalculationResults([]);
    }
  };

  if (isLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-50 via-blue-50 to-indigo-100">
        <div className="text-center">
          <div className="relative">
            <div className="animate-spin rounded-full h-16 w-16 border-4 border-blue-200 border-t-blue-600 mx-auto mb-4"></div>
            <FileSpreadsheet className="absolute top-1/2 left-1/2 transform -translate-x-1/2 -translate-y-1/2 h-6 w-6 text-blue-600" />
          </div>
          <p className="text-gray-600 font-medium">Loading Excel data...</p>
        </div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-red-50 to-pink-100">
        <div className="text-center p-8 bg-white rounded-2xl shadow-2xl border border-red-200 max-w-md mx-auto">
          <div className="text-6xl mb-4">⚠️</div>
          <h2 className="text-2xl font-bold text-red-700 mb-3">Loading Error</h2>
          <p className="text-red-600 font-medium mb-4">Failed to load Excel file</p>
          <p className="text-gray-600 text-sm mb-6">Please check if the file exists and try refreshing the page.</p>
          <button 
            onClick={() => window.location.reload()}
            className="bg-red-600 hover:bg-red-700 text-white px-6 py-2 rounded-lg font-medium transition-colors"
          >
            🔄 Retry
          </button>
        </div>
      </div>
    );
  }

  return (
    <Auth>
      <div className="min-h-screen bg-gradient-to-br from-slate-50 via-blue-50 to-indigo-100">
        <div className="container mx-auto px-4 py-8">
          <div className="bg-white rounded-2xl shadow-2xl overflow-hidden border border-gray-100">
        
          {excelData && Object.keys(excelData).length > 0 && (
            <>
              {/* Tabs */}
              <div className="border-b border-gray-200 bg-gray-50">
                <nav className="flex space-x-8 px-8">
                  {Object.keys(excelData).map((sheetName) => (
                    <button
                      key={sheetName}
                      onClick={() => {
                        setActiveTab(sheetName);
                        setCurrentPage(1);
                        setSearchTerm('');
                        setCalculationResults([]);
                        setSelectedColor('');
                      }}
                      className={`py-4 px-6 border-b-3 font-semibold text-sm transition-all duration-200 ${
                        activeTab === sheetName
                          ? 'border-indigo-500 text-indigo-600 bg-white rounded-t-lg'
                          : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                      }`}
                    >
                      {sheetName}
                    </button>
                  ))}
                  <button
                    onClick={() => {
                      setActiveTab('Calculator');
                      setCurrentPage(1);
                      setSearchTerm('');
                    }}
                    className={`py-4 px-6 border-b-3 font-semibold text-sm transition-all duration-200 ${
                      activeTab === 'Calculator'
                        ? 'border-green-500 text-green-600 bg-white rounded-t-lg'
                        : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                    }`}
                  >
                    🧮 Calculator
                  </button>
                </nav>
              </div>

              {/* Calculator Tab Content */}
              {activeTab === 'Calculator' && (
                <div className="p-8">
                  <div className="max-w-4xl mx-auto">
                    <div className="text-center mb-8">
                      <h2 className="text-3xl font-bold text-gray-800 mb-4">🧮 BOM Formulation Calculator</h2>
                      <p className="text-gray-600">Calculate precise 100kg formulations for any color from your BOM data</p>
                    </div>

                    {/* Recent Colors Quick Access */}
                    {recentColors.length > 0 && (
                      <div className="bg-gradient-to-r from-blue-50 to-indigo-50 rounded-xl p-6 mb-6 border border-blue-200">
                        <h3 className="text-lg font-semibold text-blue-800 mb-3 flex items-center gap-2">
                          🕒 Recent Colors
                        </h3>
                        <div className="flex flex-wrap gap-2">
                          {recentColors.map((colorCode, index) => {
                            const colorItem = bomHeaderColors.find(c => c.code === colorCode);
                            return colorItem ? (
                              <button
                                key={index}
                                onClick={() => {
                                  setSelectedColor(colorCode);
                                  setColorSearchTerm(`${colorCode} - ${colorItem.color}`);
                                  calculateFormulation(colorCode);
                                }}
                                className="bg-white hover:bg-blue-50 border border-blue-300 rounded-lg px-3 py-2 text-sm font-medium text-blue-700 transition-colors shadow-sm"
                              >
                                {colorCode}
                              </button>
                            ) : null;
                          })}
                        </div>
                      </div>
                    )}

                    {/* Calculator Interface */}
                    <div className="bg-gradient-to-br from-green-50 to-emerald-50 rounded-2xl p-8 mb-8 border border-green-200 shadow-lg">
                      <div className="grid md:grid-cols-2 gap-8 items-center">
                        <div>
                          <h3 className="text-xl font-semibold text-green-800 mb-4 flex items-center gap-2">
                            🎨 Select Color Formula
                            <button
                              onClick={() => {
                                if (selectedColor && !favorites.includes(selectedColor)) {
                                  setFavorites(prev => [...prev, selectedColor]);
                                }
                              }}
                              className={`ml-2 p-1 rounded-full transition-colors ${
                                selectedColor && favorites.includes(selectedColor)
                                  ? 'text-yellow-500 hover:text-yellow-600'
                                  : 'text-gray-400 hover:text-yellow-500'
                              }`}
                              title="Add to favorites"
                            >
                              ⭐
                            </button>
                          </h3>
                          <label className="block text-sm font-medium text-gray-700 mb-3">Choose from BOM Header:</label>
                          
                          {/* Searchable Dropdown */}
                          <div className="relative">
                            <input
                              type="text"
                              placeholder="🔍 Search colors..."
                              value={colorSearchTerm}
                              onChange={(e) => {
                                setColorSearchTerm(e.target.value);
                                setIsColorDropdownOpen(true);
                              }}
                              onFocus={() => setIsColorDropdownOpen(true)}
                              className="w-full px-4 text-black py-3 border border-green-300 rounded-xl focus:ring-2 focus:ring-green-500 focus:border-transparent text-lg"
                            />
                            
                            {/* Dropdown Options */}
                            {isColorDropdownOpen && (
                              <div className="absolute z-50 w-full mt-1 bg-white border border-green-300 rounded-xl shadow-lg max-h-60 overflow-y-auto">
                                {filteredColors.length > 0 ? (
                                  filteredColors.map((item, index) => (
                                    <div
                                      key={index}
                                      onClick={() => {
                                        setIsCalculating(true);
                                        setSelectedColor(item.code);
                                        setColorSearchTerm(`${item.code} - ${item.color}`);
                                        setIsColorDropdownOpen(false);
                                        
                                        // Add to recent colors
                                        setRecentColors(prev => {
                                          const updated = [item.code, ...prev.filter(c => c !== item.code)];
                                          return updated.slice(0, 5); // Keep only 5 recent
                                        });
                                        
                                        // Simulate loading for smooth UX
                                        setTimeout(() => {
                                          calculateFormulation(item.code);
                                          setIsCalculating(false);
                                        }, 100);
                                      }}
                                      className="px-4 py-3 hover:bg-green-50 cursor-pointer border-b border-gray-100 last:border-b-0"
                                    >
                                      <div className="font-medium text-gray-900">{item.code}</div>
                                      <div className="text-sm text-gray-600">{item.color}</div>
                                    </div>
                                  ))
                                ) : (
                                  <div className="px-4 py-3 text-gray-500 text-center">
                                    No colors found matching &quot;{colorSearchTerm}&quot;
                                  </div>
                                )}
                              </div>
                            )}
                          </div>
                          
                          {/* Click outside to close dropdown */}
                          {isColorDropdownOpen && (
                            <div 
                              className="fixed inset-0 z-5" 
                              onClick={() => setIsColorDropdownOpen(false)}
                            />
                          )}
                        </div>
                        
                        {selectedColor && (
                          <div className="bg-white rounded-xl p-6 border-2 border-green-300 shadow-lg relative overflow-hidden">
                            <div className="absolute top-0 right-0 w-20 h-20 bg-gradient-to-br from-green-200 to-emerald-200 rounded-full -mr-10 -mt-10 opacity-50"></div>
                            <div className="text-center relative z-10">
                              <div className="text-4xl mb-2">⚖️</div>
                              <h4 className="text-lg font-bold text-green-700">Target Weight</h4>
                              <div className="text-3xl font-bold text-green-600 mt-2">100.000 kg</div>
                              <p className="text-sm text-green-600 mt-2">Proportionally calculated</p>
                              {isCalculating && (
                                <div className="mt-3">
                                  <div className="animate-spin rounded-full h-6 w-6 border-2 border-green-300 border-t-green-600 mx-auto"></div>
                                  <p className="text-xs text-green-500 mt-1">Calculating...</p>
                                </div>
                              )}
                            </div>
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Calculation Results - Multiple Tables in Grid */}
                    {calculationResults.length > 0 && !isCalculating && (
                      <div className="animate-fadeIn">
                        {/* Header */}
                        <div className="bg-gradient-to-r from-green-600 via-emerald-600 to-teal-600 text-white p-6 rounded-t-2xl relative overflow-hidden mb-6">
                          <div className="absolute top-0 right-0 w-32 h-32 bg-white opacity-10 rounded-full -mr-16 -mt-16"></div>
                          <div className="relative z-10">
                            <h3 className="text-2xl font-bold flex items-center gap-3">
                              🎯 Formulation Results for {selectedColor}
                              <span className="bg-white bg-opacity-20 px-3 py-1 rounded-full text-sm font-medium">
                                {calculationResults.length} formulations
                              </span>
                            </h3>
                            <p className="text-green-100 mt-2">
                              {bomHeaderColors.find(item => item.code === selectedColor)?.color || 'Color'} - Multiple formulation breakdowns for 100kg production
                            </p>
                          </div>
                        </div>

                        {/* Grid of Formulation Tables */}
                        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 2xl:grid-cols-5 gap-6 mb-6">
                          {calculationResults.map((formulation, formulationIndex) => (
                            <div key={formulationIndex} className="bg-white rounded-2xl shadow-xl border border-gray-200 overflow-hidden">
                              <div className={`p-4 text-white font-bold text-center ${
                                formulationIndex === 0 ? 'bg-gradient-to-r from-blue-500 to-blue-600' :
                                formulationIndex === 1 ? 'bg-gradient-to-r from-purple-500 to-purple-600' :
                                formulationIndex === 2 ? 'bg-gradient-to-r from-orange-500 to-orange-600' :
                                formulationIndex === 3 ? 'bg-gradient-to-r from-red-500 to-red-600' :
                                formulationIndex === 4 ? 'bg-gradient-to-r from-green-500 to-green-600' :
                                formulationIndex === 5 ? 'bg-gradient-to-r from-yellow-500 to-yellow-600' :
                                formulationIndex === 6 ? 'bg-gradient-to-r from-pink-500 to-pink-600' :
                                formulationIndex === 7 ? 'bg-gradient-to-r from-indigo-500 to-indigo-600' :
                                formulationIndex === 8 ? 'bg-gradient-to-r from-teal-500 to-teal-600' :
                                formulationIndex === 9 ? 'bg-gradient-to-r from-cyan-500 to-cyan-600' :
                                'bg-gradient-to-r from-gray-500 to-gray-600'
                              }`}>
                                <h4 className="text-lg">📋 Formulation {formulationIndex + 1}</h4>
                                <p className="text-sm opacity-90">{formulation.length} ingredients</p>
                              </div>
                              
                              <div className="p-4">
                                <div className="overflow-x-auto">
                                  <table className="w-full text-sm">
                                    <thead className="bg-gray-50">
                                      <tr>
                                        <th className="px-3 py-2 text-left text-xs font-bold text-gray-700 uppercase">Code</th>
                                        <th className="px-3 py-2 text-left text-xs font-bold text-gray-700 uppercase">Description</th>
                                        <th className="px-3 py-2 text-left text-xs font-bold text-gray-700 uppercase">Weight</th>
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-gray-200">
                                      {formulation.map((row, rowIndex) => (
                                        <tr key={rowIndex} className="hover:bg-gray-50">
                                          <td className="px-3 py-2 font-medium text-gray-900">{String(row[0] || 'N/A')}</td>
                                          <td className="px-3 py-2 text-gray-700">{String(row[3] || 'N/A')}</td>
                                          <td className="px-3 py-2 font-bold text-green-600">{String(row[4] || 'N/A')} kg</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </div>
                                
                                <div className="mt-4 p-3 bg-green-50 rounded-lg border border-green-200">
                                  <div className="flex justify-between items-center">
                                    <span className="text-sm font-medium text-green-800">Total Weight:</span>
                                    <span className="text-lg font-bold text-green-600">100.000 kg</span>
                                  </div>
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>

                        {/* Export Controls */}
                        <div className="bg-gradient-to-r from-green-100 via-emerald-100 to-teal-100 rounded-xl p-6 border border-green-200">
                          <div className="flex items-center justify-between">
                            <div>
                              <h4 className="text-lg font-bold text-green-800 flex items-center gap-2">
                                ✅ All Formulations Complete
                              </h4>
                              <p className="text-green-600">{calculationResults.length} different formulation(s) ready for production</p>
                            </div>
                            <div className="flex items-center gap-4">
                              <div className="flex gap-2">
                                <button
                                  onClick={exportToCSV}
                                  className="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded-lg font-medium transition-all duration-200 flex items-center gap-2 shadow-lg hover:shadow-xl transform hover:-translate-y-0.5"
                                >
                                  📥 Export All CSV
                                </button>
                                <button
                                  onClick={() => {
                                    const allData = calculationResults.flat().map(row => row.join('\t')).join('\n');
                                    navigator.clipboard.writeText(allData);
                                  }}
                                  className="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded-lg font-medium transition-all duration-200 flex items-center gap-2 shadow-lg hover:shadow-xl transform hover:-translate-y-0.5"
                                  title="Copy all formulations to clipboard"
                                >
                                  📋 Copy All
                                </button>
                              </div>
                              <div className="text-right">
                                <div className="text-2xl font-bold text-green-600">Multiple Options</div>
                                <p className="text-sm text-green-600">Choose best formulation</p>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    )}

                    {!selectedColor && (
                      <div className="text-center py-12">
                        <div className="text-6xl mb-4 animate-bounce">🎨</div>
                        <h3 className="text-xl font-semibold text-gray-600 mb-2">Select a Color to Begin</h3>
                        <p className="text-gray-500 mb-6">Choose any color from the dropdown above to calculate its 100kg formulation</p>
                        
                        {/* Quick Stats */}
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 max-w-2xl mx-auto mt-8">
                          <div className="bg-white rounded-lg p-4 shadow-md border border-gray-200">
                            <div className="text-2xl font-bold text-blue-600">{bomHeaderColors.length}</div>
                            <div className="text-sm text-gray-600">Available Colors</div>
                          </div>
                          <div className="bg-white rounded-lg p-4 shadow-md border border-gray-200">
                            <div className="text-2xl font-bold text-green-600">100kg</div>
                            <div className="text-sm text-gray-600">Target Weight</div>
                          </div>
                          <div className="bg-white rounded-lg p-4 shadow-md border border-gray-200">
                            <div className="text-2xl font-bold text-purple-600">{recentColors.length}</div>
                            <div className="text-sm text-gray-600">Recent Calculations</div>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* Data Tables for BOM Header and BOM Item */}
              {activeTab !== 'Calculator' && (
                <>
                  {/* Search and Stats */}
                  <div className="p-6 bg-gray-50 border-b border-gray-200">
                    <div className="flex flex-col sm:flex-row justify-between items-center space-y-4 sm:space-y-0">
                      <div className="relative">
                        <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 h-4 w-4 text-gray-400" />
                        <input
                          type="text"
                          placeholder="Search data..."
                          value={searchTerm}
                          onChange={(e) => {
                            setSearchTerm(e.target.value);
                            setCurrentPage(1);
                          }}
                          className="pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-transparent"
                        />
                      </div>
                      <div className="text-sm text-gray-600">
                        Showing {paginatedData.length} of {currentData.length} records
                      </div>
                    </div>
                  </div>

                  {/* Table */}
                  <div className="overflow-x-auto">
                    {activeTab && headers.length > 0 && (
                      <table className="min-w-full divide-y divide-gray-200">
                        <thead className="bg-gradient-to-r from-gray-50 to-gray-100">
                          <tr>
                            {headers.map((header: CellValue, index: number) => (
                              <th
                                key={index}
                                className="px-6 py-4 text-left text-xs font-bold text-gray-700 uppercase tracking-wider border-r border-gray-200 last:border-r-0"
                              >
                                {String(header || '')}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody className="bg-white divide-y divide-gray-200">
                          {paginatedData.map((row: ExcelRow, rowIndex: number) => (
                            <tr key={rowIndex} className={`hover:bg-blue-50 transition-colors duration-150 ${
                              String(row[0]).includes('--- Formulation') ? 'bg-yellow-100 border-t-2 border-yellow-400' : ''
                            }`}>
                              {row.map((cell: CellValue, cellIndex: number) => (
                                <td
                                  key={cellIndex}
                                  className="px-6 py-4 whitespace-nowrap text-sm text-gray-900 border-r border-gray-100 last:border-r-0"
                                >
                                  <span className={`${cell && String(cell) !== '' ? 'text-gray-900' : 'text-gray-400 italic'} ${
                                    String(row[0]).includes('--- Formulation') ? 'font-bold text-yellow-800' : ''
                                  }`}>
                                    {cell && String(cell) !== '' ? String(cell) : 'N/A'}
                                  </span>
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                </>
              )}

              {/* Pagination */}
              {totalPages > 1 && (
                <div className="bg-gray-50 px-6 py-4 border-t border-gray-200">
                  <div className="flex items-center justify-between">
                    <div className="text-sm text-gray-700">
                      Page {currentPage} of {totalPages}
                    </div>
                    <div className="flex space-x-2">
                      <button
                        onClick={() => setCurrentPage(Math.max(1, currentPage - 1))}
                        disabled={currentPage === 1}
                        className="flex items-center px-3 py-2 text-sm font-medium text-gray-500 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                      >
                        <ChevronLeft className="h-4 w-4 mr-1" />
                        Previous
                      </button>
                      
                      {/* Page numbers */}
                      <div className="flex space-x-1">
                        {Array.from({ length: Math.min(5, totalPages) }, (_, i) => {
                          const pageNum = Math.max(1, Math.min(totalPages - 4, currentPage - 2)) + i;
                          return (
                            <button
                              key={pageNum}
                              onClick={() => setCurrentPage(pageNum)}
                              className={`px-3 py-2 text-sm font-medium rounded-lg transition-colors ${
                                currentPage === pageNum
                                  ? 'bg-indigo-600 text-white'
                                  : 'text-gray-500 bg-white border border-gray-300 hover:bg-gray-50'
                              }`}
                            >
                              {pageNum}
                            </button>
                          );
                        })}
                      </div>

                      <button
                        onClick={() => setCurrentPage(Math.min(totalPages, currentPage + 1))}
                        disabled={currentPage === totalPages}
                        className="flex items-center px-3 py-2 text-sm font-medium text-gray-500 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                      >
                        Next
                        <ChevronRight className="h-4 w-4 ml-1" />
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </>
          )}

          {(!excelData || Object.keys(excelData).length === 0) && !isLoading && (
            <div className="p-12 text-center">
              <div className="bg-gradient-to-br from-gray-50 to-gray-100 rounded-2xl p-8 max-w-md mx-auto">
                <FileSpreadsheet className="h-16 w-16 text-gray-300 mx-auto mb-4" />
                <h3 className="text-xl font-semibold text-gray-700 mb-2">No Data Found</h3>
                <p className="text-gray-500 text-lg mb-4">The Excel file appears to be empty or corrupted.</p>
                <div className="text-sm text-gray-400">
                  <p>Expected file: <code className="bg-gray-200 px-2 py-1 rounded">Merge Source Data of BOM Header & BOM Item.xlsx</code></p>
                </div>
              </div>
            </div>
          )}
        </div>
        </div>
      </div>
    </Auth>
  );
}
