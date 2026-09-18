import React, { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { 
  Chart as ChartJS, 
  CategoryScale, 
  LinearScale, 
  RadialLinearScale,
  BarElement, 
  Title, 
  Tooltip, 
  Legend, 
  PointElement, 
  LineElement, 
  ArcElement 
} from 'chart.js';
import { Bar, Line, Pie, Doughnut, Radar, PolarArea } from 'react-chartjs-2';
import { saveAs } from 'file-saver';
import ThreeDChart from '../components/ThreeDChart';
import DashboardLayout from '../components/Dashboard/Layout';
import activityTracker from '../utils/activityTracker';

// Register Chart.js components
ChartJS.register(
  CategoryScale,
  LinearScale,
  RadialLinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
  PointElement,
  LineElement,
  ArcElement
);

/**
 * UserDashboard Component
 * Main dashboard for regular users to upload Excel files and configure chart settings
 * Features: File upload, Excel parsing, chart type selection, data mapping, chart generation
 */
const UserDashboard = () => {
  const [jobs, setJobs] =useState([]);
  const [selectedJob, setSelectedJob] = useState(null);
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [selectedFile, setSelectedFile] = useState(null);
  const [chartType, setChartType] = useState('2d'); // '2d' or '3d'
  const [specificChart, setSpecificChart] = useState('bar'); // Chart type selection
  const [xAxis, setXAxis] = useState(''); // X-axis data selection
  const [yAxis, setYAxis] = useState(''); // Y-axis data selection
  const [isDragOver, setIsDragOver] = useState(false); // Drag and drop state
  const [excelData, setExcelData] = useState(null); // Parsed Excel data
  const [sheetNames, setSheetNames] = useState([]); // Available sheets
  const [selectedSheet, setSelectedSheet] = useState(''); // Selected sheet
  const [columns, setColumns] = useState([]); // Available columns
  const [generatedChart, setGeneratedChart] = useState(null); // Generated chart data
  const [isDownloading, setIsDownloading] = useState(false); // Download state
  const chartRef = useRef(null); // Reference to current chart instance
  const threeRendererRef = useRef(null); // Reference to Three.js renderer
  const [maintenance, setMaintenance] = useState(false); // System maintenance state
  const [maintenanceUntil, setMaintenanceUntil] = useState(null); // Timestamp (ms) when maintenance ends
  const [maintenanceStarted, setMaintenanceStarted] = useState(null); // Timestamp (ms) when maintenance started
  const [countdown, setCountdown] = useState({ days: 0, hours: 0, minutes: 0, seconds: 0 });
  const [progressPct, setProgressPct] = useState(0);

//fetch user job from server
const fetchUserJobs = async () => {
  try{
    const response = await fetch('/api/jobs');
    if(response.ok){
      const data = await response.json();
      setJobs(data);
    }
  }
  catch (err) {
    console.error('Error fetching jobs',err);
  }
}

//poll jobs Every 3 second for live updates
useEffect(() =>{
  fetchUserJobs();
  const interval = setInterval(fetchUserJobs,3000);
  return () => clearInterval(interval);
},[]);

  // Read maintenance flag and react to changes from System Settings (top-level hook)
  useEffect(() => {
    // Log dashboard access
    const user = JSON.parse(localStorage.getItem('user') || 'null');
    activityTracker.logActivity('dashboard_access', user?.email || 'anonymous');

    try {
      setMaintenance(!!JSON.parse(localStorage.getItem('sys_maintenance') || 'false'));
      const until = JSON.parse(localStorage.getItem('sys_maintenance_until') || 'null');
      setMaintenanceUntil(typeof until === 'number' ? until : null);
      const started = JSON.parse(localStorage.getItem('sys_maintenance_started') || 'null');
      setMaintenanceStarted(typeof started === 'number' ? started : null);
    } catch {}
    const onStorage = (e) => {
      if (e.key === 'sys_maintenance') {
        try { setMaintenance(!!JSON.parse(e.newValue || 'false')); } catch {}
      }
      if (e.key === 'sys_maintenance_until') {
        try { setMaintenanceUntil(e.newValue ? JSON.parse(e.newValue) : null); } catch {}
      }
      if (e.key === 'sys_maintenance_started') {
        try { setMaintenanceStarted(e.newValue ? JSON.parse(e.newValue) : null); } catch {}
      }
    };
    window.addEventListener('storage', onStorage);
    return () => window.removeEventListener('storage', onStorage);
  }, []);

  // Countdown and auto-expire logic
  useEffect(() => {
    const tick = () => {
      if (!maintenanceUntil) { setCountdown({ days: 0, hours: 0, minutes: 0, seconds: 0 }); return; }
      const now = Date.now();
      const ms = Math.max(0, maintenanceUntil - now);
      if (ms === 0 && maintenance) {
        // Auto-end maintenance
        try {
          localStorage.setItem('sys_maintenance', JSON.stringify(false));
          localStorage.removeItem('sys_maintenance_until');
          localStorage.removeItem('sys_maintenance_started');
        } catch {}
        setMaintenance(false);
        setMaintenanceUntil(null);
        setMaintenanceStarted(null);
        setCountdown({ days: 0, hours: 0, minutes: 0, seconds: 0 });
        setProgressPct(100);
        return;
      }
      const DAY = 24*60*60*1000, HOUR = 60*60*1000, MIN = 60*1000;
      const days = Math.floor(ms / DAY);
      const hours = Math.floor((ms % DAY) / HOUR);
      const minutes = Math.floor((ms % HOUR) / MIN);
      const seconds = Math.floor((ms % MIN) / 1000);
      setCountdown({ days, hours, minutes, seconds });

      // Progress percent (elapsed/total)
      if (maintenanceStarted && maintenanceUntil && maintenanceUntil > maintenanceStarted) {
        const total = maintenanceUntil - maintenanceStarted;
        const elapsed = Math.min(total, Math.max(0, Date.now() - maintenanceStarted));
        const pct = Math.max(0, Math.min(100, (elapsed / total) * 100));
        setProgressPct(pct);
      } else {
        setProgressPct(0);
      }
    };
    tick();
    const id = setInterval(tick, 1000);
    return () => clearInterval(id);
  }, [maintenanceUntil, maintenance, maintenanceStarted]);

  /**
   * Handle file upload and Excel parsing
   * @param {File} file - The uploaded Excel file
   */
  const handleFileUpload = async(file) => {
    const user = JSON.parse(localStorage.getItem('user') || 'null');
    // System settings enforcement
    try {
      const maintenance = !!JSON.parse(localStorage.getItem('sys_maintenance') || 'false');
      const uploadLimitMB = Number(JSON.parse(localStorage.getItem('sys_upload_limit') || '0')) || 0;
      if (maintenance) {
        alert('Uploads are disabled while maintenance mode is ON.');
        return;
      }
      if (uploadLimitMB > 0) {
        const maxBytes = uploadLimitMB * 1024 * 1024;
        if (file && file.size > maxBytes) {
          alert(`File exceeds the upload limit of ${uploadLimitMB} MB.`);
          return;
        }
      }
    } catch {}
    const formdata = new FormData();
    formData.append('file',file);

    try{
      const res = await fetch('/api/jobs/upload',{
        method:'POST',
        body:formData
      });
      const data = await res.json();

      if(res.ok){
        alert('File submitted successfully');
        fetchUserJobs();
        activityTracker.logActivity('file_upload', user?.email || 'ananymous',{
          fileName:file.name,
          jobId:data.jobId
        });
      }else{
        alert(data.error || 'upload failed');
      }
    } catch (err){
      console.error('Upload Error:', err);
      alert('FAiled to upload file to backend server');
    }
  };

  //load parsed Background job output into state

  const handleLoadJobResult = (job) =>{
    if(job.status !=='COMPLETED' || !job.result || !job.result.data){
      alert('job results are not ready yet');
      return;
    }
    const rows = jobs.reslut.data;
    if(rows.length ===0){
      alert('Uploaded file contains no data row');
      return;
    }
    const headers = Object.keys(rows[0]);
    const sheetData = [headers, ...rows.map(row =>Object.values(row))];

    setColumns(headers);
    setSelectedJob(job);
    setExcelData({
      headers,
      sheets:sheetData
    });
    console.log('job results loaded into view', job._id);
  }
  const cleanupChart = () =>{
    if(chartRef.current){
      chartRef.current.destroy();
      chartRef.current =null;
    }
  }

    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        // Parse Excel file with enhanced options for Excel 2024 compatibility
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellDates: true, // Parse dates properly
          cellNF: false,   // Don't parse number formats
          cellStyles: false, // Don't parse cell styles for performance
          cellText: false   // Don't parse text formatting
        });
        
        // Get sheet names
        const sheets = workbook.SheetNames;
        setSheetNames(sheets);
        setSelectedSheet(sheets[0]); // Select first sheet by default
        
        // Parse first sheet data
        const firstSheet = workbook.Sheets[sheets[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
        
        // Extract headers (first row)
        const headers = jsonData[0] || [];
        setColumns(headers);
        
        // Store parsed data
        setExcelData({
          workbook,
          sheets: jsonData,
          headers
        });
        
        // Update file list
        const timestamp = new Date().toISOString();
        const newFile = {
          name: file.name,
          size: file.size,
          date: new Date().toLocaleDateString(),
          charts: 0,
          downloads: 0,
          timestamp
        };

        setUploadedFiles(prev => [newFile, ...prev]);
        setSelectedFile(newFile);

        // Store upload in history
        const uploadRecord = {
          name: file.name,
          size: file.size,
          timestamp: new Date().toISOString(),
          user: user?.email || 'anonymous'
        };
        const history = JSON.parse(localStorage.getItem('upload_history') || '[]');
        history.push(uploadRecord);
        try {
          localStorage.setItem('upload_history', JSON.stringify(history));
        } catch (e) {
          console.warn('Could not write upload_history', e);
        }

        // Track upload activity
        activityTracker.logActivity('file_upload', user?.email || 'anonymous', {
          fileName: file.name,
          fileSize: file.size,
          fileType: file.name.split('.').pop()?.toLowerCase()
        });

        
        console.log('✅ Excel file parsed successfully:', {
          fileName: file.name,
          fileType: file.type,
          sheets: sheets,
          headers: headers,
          rowCount: jsonData.length - 1,
          excelVersion: 'Excel 2024 Compatible'
        });
        
      } catch (error) {
        console.error('❌ Error parsing Excel file:', error);
        alert('Error parsing Excel file. Please ensure it\'s a valid Excel file.');
      }
    };
    
    reader.readAsArrayBuffer(file);
  };

  /**
   * Handle sheet selection change
   * @param {string} sheetName - Selected sheet name
   */
  const handleSheetChange = (sheetName) => {
    if (!excelData) return;
    
    try {
      const sheet = excelData.workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      const headers = jsonData[0] || [];
      
      setSelectedSheet(sheetName);
      setColumns(headers);
      setExcelData(prev => ({
        ...prev,
        sheets: jsonData,
        headers
      }));
      
      // Reset axes when sheet changes
      setXAxis('');
      setYAxis('');
      
      console.log('📊 Sheet changed to:', sheetName, 'Headers:', headers);
    } catch (error) {
      console.error('❌ Error changing sheet:', error);
    }
  };

  /**
   * Clean up existing chart instance
   */
  const cleanupChart = () => {
    if (chartRef.current) {
      chartRef.current.destroy();
      chartRef.current = null;
    }
  };

  /**
   * Generate chart data based on selected configuration
   */
  const generateChartData = () => {
    if (!excelData || !xAxis || !yAxis) {
      alert('Please select both X and Y axis data');
      return;
    }

    // Track chart generation activity
    const user = JSON.parse(localStorage.getItem('user') || 'null');
    activityTracker.logActivity('chart_generation', user?.email || 'anonymous', {
      chartType: chartType,
      xAxis: xAxis,
      yAxis: yAxis,
      dataPoints: excelData.length
    });
    
    try {
      const data = excelData.sheets.slice(1); // Skip header row
      const xIndex = excelData.headers.indexOf(xAxis);
      const yIndex = excelData.headers.indexOf(yAxis);
      
      if (xIndex === -1 || yIndex === -1) {
        alert('Selected columns not found in data');
        return;
      }
      
      // Extract data for chart
      const labels = data.map(row => row[xIndex]).filter(Boolean);
      const values = data.map(row => parseFloat(row[yIndex]) || 0);
      
      // Validate data for specific chart types
      if (['pie', 'doughnut', 'polar'].includes(specificChart)) {
        if (values.length === 0 || values.every(v => v === 0)) {
          alert('Pie/Doughnut/Polar charts need numeric values greater than 0');
          return;
        }
      }
      
      // Create chart data based on chart type
      let chartData;
      
      if (['pie', 'doughnut', 'polar'].includes(specificChart)) {
        // Pie, Doughnut, and Polar charts need different data structure
        chartData = {
          labels: labels,
          datasets: [{
            data: values,
            backgroundColor: [
              'rgba(255, 99, 132, 0.8)',
              'rgba(54, 162, 235, 0.8)',
              'rgba(255, 206, 86, 0.8)',
              'rgba(75, 192, 192, 0.8)',
              'rgba(153, 102, 255, 0.8)',
              'rgba(255, 159, 64, 0.8)',
              'rgba(199, 199, 199, 0.8)',
              'rgba(83, 102, 255, 0.8)',
            ],
            borderColor: [
              'rgba(255, 99, 132, 1)',
              'rgba(54, 162, 235, 1)',
              'rgba(255, 206, 86, 1)',
              'rgba(75, 192, 192, 1)',
              'rgba(153, 102, 255, 1)',
              'rgba(255, 159, 64, 1)',
              'rgba(199, 199, 199, 1)',
              'rgba(83, 102, 255, 1)',
            ],
            borderWidth: 2
          }]
        };
      } else {
        // Bar, Line, Radar charts use standard structure
        chartData = {
          labels: labels,
          datasets: [{
            label: yAxis,
            data: values,
            backgroundColor: [
              'rgba(255, 99, 132, 0.8)',
              'rgba(54, 162, 235, 0.8)',
              'rgba(255, 206, 86, 0.8)',
              'rgba(75, 192, 192, 0.8)',
              'rgba(153, 102, 255, 0.8)',
            ],
            borderColor: [
              'rgba(255, 99, 132, 1)',
              'rgba(54, 162, 235, 1)',
              'rgba(255, 206, 86, 1)',
              'rgba(75, 192, 192, 1)',
              'rgba(153, 102, 255, 1)',
            ],
            borderWidth: 2
          }]
        };
      }
      
      // Clean up existing chart before creating new one
      cleanupChart();
      
      setGeneratedChart(chartData);
      
      // Update file stats
      setUploadedFiles(prev => 
        prev.map(f => 
          f.name === selectedFile.name 
            ? { ...f, charts: f.charts + 1 }
            : f
        )
      );
      
      console.log('📊 Chart generated:', {
        type: specificChart,
        dimension: chartType,
        xAxis,
        yAxis,
        dataPoints: labels.length
      });
      
    } catch (error) {
      console.error('❌ Error generating chart:', error);
      alert('Error generating chart. Please check your data selection.');
    }
  };

  /**
   * Download chart as image (PNG)
   */
  const downloadChartAsImage = () => {
    if (!generatedChart || isDownloading) return;
    
    setIsDownloading(true);
    
    try {
      // For 3D charts, we need to capture the Three.js canvas
      if (chartType === '3d') {
        // Wait for the canvas to be fully rendered
        setTimeout(() => {
          if (threeRendererRef.current) {
            const { gl, scene, camera } = threeRendererRef.current;
            
            // Force a render
            gl.render(scene, camera);
            
            // Get the canvas from the renderer
            const canvas = gl.domElement;
            
            if (canvas && canvas.width > 0 && canvas.height > 0) {
              console.log('Three.js canvas found:', canvas.width, 'x', canvas.height);
              
              // Create a temporary canvas with proper dimensions
              const tempCanvas = document.createElement('canvas');
              const tempCtx = tempCanvas.getContext('2d');
              
              // Set proper dimensions
              tempCanvas.width = canvas.width;
              tempCanvas.height = canvas.height;
              
              // Fill with background color first
              tempCtx.fillStyle = '#1a1a2e';
              tempCtx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);
              
              // Copy the Three.js canvas content
              tempCtx.drawImage(canvas, 0, 0);
              
              // Convert to blob with high quality
              tempCanvas.toBlob((blob) => {
                setIsDownloading(false);
                if (blob && blob.size > 1000) { // Check if blob has content
                  saveAs(blob, `3d-chart-${specificChart}-${Date.now()}.png`);
                  
                  // Update download count
                  setUploadedFiles(prev => 
                    prev.map(f => 
                      f.name === selectedFile.name 
                        ? { ...f, downloads: f.downloads + 1 }
                        : f
                    )
                  );
                  
                  console.log('✅ 3D chart downloaded successfully');
                } else {
                  console.error('❌ Blob is empty or too small:', blob?.size);
                  alert('Failed to capture chart. Please try again.');
                }
              }, 'image/png', 1.0);
            } else {
              setIsDownloading(false);
              console.error('❌ Three.js canvas not found or invalid dimensions');
              alert('Chart canvas not found. Please try again.');
            }
          } else {
            setIsDownloading(false);
            console.error('❌ Three.js renderer not available');
            alert('3D chart not ready. Please try again.');
          }
        }, 2000); // Wait 2 seconds for rendering
      } else {
        // For 2D charts, use the chart canvas
        setTimeout(() => {
          // Try multiple ways to find the chart canvas
          const canvas = document.getElementById('chart-canvas') || 
                        document.querySelector('canvas[data-chartjs-chart]') ||
                        document.querySelector('canvas');
          
          if (canvas && canvas.width > 0 && canvas.height > 0) {
            console.log('2D chart canvas found:', canvas.width, 'x', canvas.height);
            
            // Create a temporary canvas for better quality
            const tempCanvas = document.createElement('canvas');
            const tempCtx = tempCanvas.getContext('2d');
            tempCanvas.width = canvas.width;
            tempCanvas.height = canvas.height;
            
            // Fill with background color first
            tempCtx.fillStyle = '#1a1a2e';
            tempCtx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);
            
            // Copy the chart canvas content
            tempCtx.drawImage(canvas, 0, 0);
            
            tempCanvas.toBlob((blob) => {
              setIsDownloading(false);
              if (blob && blob.size > 1000) {
                saveAs(blob, `chart-${specificChart}-${Date.now()}.png`);
                
                // Update download count
                setUploadedFiles(prev => 
                  prev.map(f => 
                    f.name === selectedFile.name 
                      ? { ...f, downloads: f.downloads + 1 }
                      : f
                  )
                );
                
                console.log('✅ 2D chart downloaded successfully');
              } else {
                console.error('❌ 2D chart blob is empty or too small:', blob?.size);
                alert('Failed to capture 2D chart. Please try again.');
              }
            }, 'image/png', 1.0);
          } else {
            setIsDownloading(false);
            console.error('❌ 2D chart canvas not found or invalid dimensions');
            alert('2D chart canvas not found. Please try again.');
          }
        }, 1000); // Wait 1 second for 2D chart rendering
      }
    } catch (error) {
      setIsDownloading(false);
      console.error('❌ Error downloading chart:', error);
      alert('Error downloading chart. Please try again.');
    }
  };

  /**
   * Download chart as PDF
   */
  const downloadChartAsPDF = () => {
    if (!generatedChart || isDownloading) return;
    
    setIsDownloading(true);
    
    try {
      // For 3D charts, we need to capture the Three.js canvas
      if (chartType === '3d') {
        setTimeout(() => {
          if (threeRendererRef.current) {
            const { gl, scene, camera } = threeRendererRef.current;
            
            // Force a render
            gl.render(scene, camera);
            
            // Get the canvas from the renderer
            const canvas = gl.domElement;
            
            if (canvas && canvas.width > 0 && canvas.height > 0) {
              console.log('Three.js canvas found for PDF:', canvas.width, 'x', canvas.height);
              
              // Create a temporary canvas with proper dimensions
              const tempCanvas = document.createElement('canvas');
              const tempCtx = tempCanvas.getContext('2d');
              tempCanvas.width = canvas.width;
              tempCanvas.height = canvas.height;
              
              // Fill with background color first
              tempCtx.fillStyle = '#1a1a2e';
              tempCtx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);
              
              // Copy the Three.js canvas content
              tempCtx.drawImage(canvas, 0, 0);
              
              // Convert to blob and create PDF
              tempCanvas.toBlob((blob) => {
                setIsDownloading(false);
                if (blob && blob.size > 1000) {
                  // Create PDF from image
                  createPDFFromImage(blob, `3D ${specificChart.toUpperCase()} Chart`);
                  
                  // Update download count
                  setUploadedFiles(prev => 
                    prev.map(f => 
                      f.name === selectedFile.name 
                        ? { ...f, downloads: f.downloads + 1 }
                        : f
                    )
                  );
                  
                  console.log('✅ 3D chart PDF downloaded successfully');
                } else {
                  console.error('❌ Blob is empty or too small for PDF:', blob?.size);
                  alert('Failed to capture chart for PDF. Please try again.');
                }
              }, 'image/png', 1.0);
            } else {
              setIsDownloading(false);
              console.error('❌ Three.js canvas not found for PDF');
              alert('Chart canvas not found for PDF. Please try again.');
            }
          } else {
            setIsDownloading(false);
            console.error('❌ Three.js renderer not available for PDF');
            alert('3D chart not ready for PDF. Please try again.');
          }
        }, 2000);
      } else {
        // For 2D charts, use the chart canvas
        setTimeout(() => {
          const canvas = document.getElementById('chart-canvas') || 
                        document.querySelector('canvas[data-chartjs-chart]') ||
                        document.querySelector('canvas');
          
          if (canvas && canvas.width > 0 && canvas.height > 0) {
            console.log('2D chart canvas found for PDF:', canvas.width, 'x', canvas.height);
            
            // Create a temporary canvas for better quality
            const tempCanvas = document.createElement('canvas');
            const tempCtx = tempCanvas.getContext('2d');
            tempCanvas.width = canvas.width;
            tempCanvas.height = canvas.height;
            
            // Fill with background color first
            tempCtx.fillStyle = '#1a1a2e';
            tempCtx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);
            
            // Copy the chart canvas content
            tempCtx.drawImage(canvas, 0, 0);
            
            tempCanvas.toBlob((blob) => {
              setIsDownloading(false);
              if (blob && blob.size > 1000) {
                // Create PDF from image
                createPDFFromImage(blob, `${specificChart.toUpperCase()} Chart`);
                
                // Update download count
                setUploadedFiles(prev => 
                  prev.map(f => 
                    f.name === selectedFile.name 
                      ? { ...f, downloads: f.downloads + 1 }
                      : f
                  )
                );
                
                console.log('✅ 2D chart PDF downloaded successfully');
              } else {
                console.error('❌ 2D chart blob is empty for PDF:', blob?.size);
                alert('Failed to capture 2D chart for PDF. Please try again.');
              }
            }, 'image/png', 1.0);
          } else {
            setIsDownloading(false);
            console.error('❌ 2D chart canvas not found for PDF');
            alert('2D chart canvas not found for PDF. Please try again.');
          }
        }, 1000);
      }
    } catch (error) {
      setIsDownloading(false);
      console.error('❌ Error downloading chart as PDF:', error);
      alert('Error downloading chart as PDF. Please try again.');
    }
  };

  /**
   * Create PDF from image blob
   */
  const createPDFFromImage = (imageBlob, title) => {
    // Create a simple PDF using canvas
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    canvas.width = 800;
    canvas.height = 600;
    
    // Fill background
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    
    // Add title
    ctx.fillStyle = '#000000';
    ctx.font = 'bold 24px Arial';
    ctx.textAlign = 'center';
    ctx.fillText(title, canvas.width / 2, 40);
    
    // Add timestamp
    ctx.font = '14px Arial';
    ctx.fillText(`Generated on: ${new Date().toLocaleString()}`, canvas.width / 2, 70);
    
    // Load and draw the chart image
    const img = new Image();
    img.onload = () => {
      // Calculate image dimensions to fit in PDF
      const maxWidth = canvas.width - 40;
      const maxHeight = canvas.height - 100;
      const scale = Math.min(maxWidth / img.width, maxHeight / img.height);
      const scaledWidth = img.width * scale;
      const scaledHeight = img.height * scale;
      const x = (canvas.width - scaledWidth) / 2;
      const y = 100;
      
      ctx.drawImage(img, x, y, scaledWidth, scaledHeight);
      
      // Convert to blob and download
      canvas.toBlob((pdfBlob) => {
        saveAs(pdfBlob, `${title.replace(/\s+/g, '-')}-${Date.now()}.pdf`);
      }, 'application/pdf');
    };
    img.src = URL.createObjectURL(imageBlob);
  };

  /**
   * Handle drag over event
   * @param {DragEvent} e - Drag event
   */
  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  /**
   * Handle drag leave event
   * @param {DragEvent} e - Drag event
   */
  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragOver(false);
  };

  /**
   * Handle drop event
   * @param {DragEvent} e - Drag event
   */
  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragOver(false);
    
    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) {
      const file = files[0];
      // Check for Excel file types including Excel 2024 formats
      const excelTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.ms-excel', // .xls
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsm
        'application/vnd.ms-excel.sheet.macroEnabled.12', // .xlsm
        'application/vnd.ms-excel.sheet.binary.macroEnabled.12', // .xlsb
        'application/vnd.openxmlformats-officedocument.spreadsheetml.template', // .xltx
        'application/vnd.ms-excel.template.macroEnabled.12' // .xltm
      ];
      
      if (excelTypes.includes(file.type) || 
          file.name.match(/\.(xlsx|xls|xlsm|xlsb|xltx|xltm)$/i)) {
        handleFileUpload(file);
      } else {
        alert('Please upload a valid Excel file (.xlsx, .xls, .xlsm, .xlsb, .xltx, .xltm)');
      }
    }
  };

  /**
   * Handle file input change
   * @param {Event} e - File input event
   */
  const handleFileInput = (e) => {
    const file = e.target.files[0];
    if (file) {
      handleFileUpload(file);
    }
  };

  // Cleanup chart on component unmount
  useEffect(() => {
    return () => {
      cleanupChart();
    };
  }, []);

  // Chart type configurations with icons and names
  const chartTypes = {
    '2d': [
      { id: 'bar', name: 'Bar Chart', icon: '📊' },
      { id: 'line', name: 'Line Chart', icon: '📈' },
      { id: 'pie', name: 'Pie Chart', icon: '🥧' },
      { id: 'doughnut', name: 'Doughnut Chart', icon: '🍩' },
      { id: 'radar', name: 'Radar Chart', icon: '🎯' },
      { id: 'polar', name: 'Polar Area', icon: '🌊' }
    ],
    '3d': [
      { id: 'bar', name: '3D Bar Chart', icon: '🏗️' },
      { id: 'line', name: '3D Line Chart', icon: '📊' },
      { id: 'pie', name: '3D Pie Chart', icon: '🥧' }
    ]
  };

  // If maintenance is ON, show only the maintenance screen
  if (maintenance) {
    return (
      <DashboardLayout title="Dashboard">
        <div className="min-h-[70vh] flex items-center justify-center px-4">
          {/* Scoped animation styles */}
          <style>{`
            @keyframes gradientMove { 0% { background-position: 0% 50% } 100% { background-position: 200% 50% } }
            @keyframes slowSpin { from { transform: rotate(0deg) } to { transform: rotate(360deg) } }
          `}</style>
          <div className="relative w-full max-w-2xl">
            {/* Aura background */}
            <div className="absolute -inset-8 blur-3xl opacity-25 bg-gradient-to-r from-rose-500 via-amber-400 to-emerald-400 rounded-[48px]" />

            {/* Glass card */}
            <div className="relative rounded-3xl border border-white/10 bg-white/10 backdrop-blur-2xl p-10 text-center shadow-2xl">
              {/* Conic gradient spinner */}
              <div className="mx-auto mb-8 relative w-28 h-28">
                <div className="absolute inset-0 rounded-full animate-[slowSpin_10s_linear_infinite] opacity-80 bg-[conic-gradient( from_0deg, rgba(244,63,94,0.0), rgba(244,63,94,0.6), rgba(245,158,11,0.0), rgba(16,185,129,0.6), rgba(99,102,241,0.0) )]" style={{ mask: 'radial-gradient(farthest-side, transparent 58%, black 59%)', WebkitMask: 'radial-gradient(farthest-side, transparent 58%, black 59%)' }} />
                <div className="absolute inset-3 rounded-full animate-[slowSpin_18s_linear_infinite] opacity-60 bg-[conic-gradient( from_0deg, rgba(16,185,129,0.0), rgba(16,185,129,0.6), rgba(14,165,233,0.0), rgba(139,92,246,0.6), rgba(244,63,94,0.0) )]" style={{ mask: 'radial-gradient(farthest-side, transparent 58%, black 59%)', WebkitMask: 'radial-gradient(farthest-side, transparent 58%, black 59%)' }} />
                <div className="absolute inset-0 flex items-center justify-center">
                  <div className="w-16 h-16 rounded-2xl bg-white/10 border border-white/10 flex items-center justify-center shadow-lg">
                    <span className="text-2xl">🛠️</span>
                  </div>
                </div>
              </div>

              <h2 className="text-4xl font-extrabold bg-gradient-to-r from-rose-300 via-amber-300 to-emerald-300 bg-clip-text text-transparent mb-3">
                We’ll be right back
              </h2>
              <p className="text-gray-200/90 max-w-xl mx-auto mb-4">
                We’re performing scheduled maintenance to keep everything fast, secure, and reliable. Thank you for your patience.
              </p>
              {maintenanceUntil && (
                <div className="mb-6 text-gray-100/90">
                  <span className="text-sm uppercase tracking-wide text-white/70">Time remaining: </span>
                  <span className="text-lg font-semibold">{countdown.days}d {countdown.hours}h {countdown.minutes}m {String(countdown.seconds).padStart(2,'0')}s</span>
                </div>
              )}

              {/* Progress bar with subtle shimmer */}
              <div className="w-full max-w-lg mx-auto">
                <div className="h-3 rounded-full bg-white/10 border border-white/10 overflow-hidden">
                  <div className="h-full rounded-full bg-[linear-gradient(90deg,rgba(244,63,94,0.8),rgba(245,158,11,0.8),rgba(16,185,129,0.8))] bg-[length:200%_100%]" style={{ width: `${progressPct}%`, animation: 'gradientMove 2.4s linear infinite' }} />
                </div>
                <div className="flex items-center justify-between mt-2 text-xs text-gray-300">
                  <span>Applying updates</span>
                  <span>{maintenanceUntil ? `${Math.round(progressPct)}%` : 'Duration not scheduled'}</span>
                </div>
              </div>

              {/* Helper links */}
              <div className="mt-8 text-sm text-gray-300">
                Experiencing issues? Check the <span className="text-emerald-300 font-semibold">status</span> page or contact <span className="text-amber-300 font-semibold">support</span>.
              </div>
            </div>
          </div>
        </div>
      </DashboardLayout>
    );
  }

  return (
    <DashboardLayout title="Dashboard">
    <div className="min-h-screen bg-transparent py-4 sm:py-8 px-2 sm:px-4 lg:px-8 relative overflow-hidden">
      
      <div className="max-w-7xl mx-auto relative z-10">
        {/* Header Section */}
        <div className="text-center mb-6 sm:mb-8">
          <div className="inline-flex items-center justify-center w-20 h-20 bg-gradient-to-r from-cyan-400 to-blue-500 rounded-full mb-4 shadow-2xl">
            <span className="text-3xl">📊</span>
          </div>
          <h1 className="text-3xl sm:text-4xl lg:text-5xl font-extrabold bg-gradient-to-r from-cyan-400 via-blue-500 to-purple-500 bg-clip-text text-transparent mb-3 sm:mb-4 drop-shadow-lg">
            Excel Analytics Dashboard
          </h1>
          <p className="text-sm sm:text-lg text-cyan-100 max-w-2xl mx-auto px-4 leading-relaxed">
            Transform your data into stunning visualizations with our powerful Excel analytics platform
          </p>
        </div>

        {/* Main Content Container */}
        <div className="backdrop-blur-xl bg-gradient-to-br from-white/10 via-white/5 to-white/10 border border-white/20 rounded-3xl shadow-2xl p-6 sm:p-8 md:p-10 animate-slide-up space-y-8 sm:space-y-10 relative overflow-hidden">
          {/* Subtle gradient overlay */}
          <div className="absolute inset-0 bg-gradient-to-br from-cyan-500/5 via-transparent to-purple-500/5 pointer-events-none"></div>
          
          {/* File Upload Section */}
          <div className="space-y-6 relative z-10">
            <h2 className="text-2xl sm:text-3xl font-bold bg-gradient-to-r from-cyan-400 to-blue-500 bg-clip-text text-transparent flex items-center gap-3">
              <div className="w-10 h-10 bg-gradient-to-r from-cyan-400 to-blue-500 rounded-lg flex items-center justify-center shadow-lg">
                <span className="text-xl">📁</span>
              </div>
              Upload Excel File
            </h2>
            
            {/* Drag & Drop Zone */}
            <div
              className={`border-2 border-dashed rounded-2xl p-8 sm:p-10 text-center transition-all duration-300 relative overflow-hidden ${
                isDragOver
                  ? 'border-cyan-400 bg-gradient-to-r from-cyan-400/20 to-blue-400/20 shadow-2xl scale-105'
                  : 'border-cyan-300/50 hover:border-cyan-400 hover:bg-gradient-to-r hover:from-cyan-400/10 hover:to-blue-400/10'
              }`}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
            >
              {/* Animated background for drag over */}
              {isDragOver && (
                <div className="absolute inset-0 bg-gradient-to-r from-cyan-400/20 to-blue-400/20 animate-pulse"></div>
              )}
              
              <div className="relative z-10">
                <div className="text-6xl sm:text-8xl mb-6 animate-bounce">📊</div>
                <h3 className="text-xl sm:text-2xl font-bold text-cyan-100 mb-4">
                  Drop your Excel file here
                </h3>
                <p className="text-cyan-200 mb-6 text-sm sm:text-base leading-relaxed">
                  Drag and drop your Excel file here, or click the button below to browse
                </p>
                <div className="bg-gradient-to-r from-cyan-500/20 to-blue-500/20 rounded-xl p-4 mb-6">
                  <p className="text-cyan-100 text-sm font-medium">
                    📋 Supports: .xlsx, .xls, .xlsm, .xlsb, .xltx, .xltm
                  </p>
                  <p className="text-cyan-200 text-xs mt-1">
                    Excel 2024 compatible with advanced features
                  </p>
                </div>
                <input
                  type="file"
                  accept=".xlsx,.xls,.xlsm,.xlsb,.xltx,.xltm"
                  onChange={handleFileInput}
                  className="hidden"
                  id="file-input"
                />
                <label
                  htmlFor="file-input"
                  className="inline-flex items-center px-8 py-4 bg-gradient-to-r from-cyan-500 to-blue-600 text-white font-bold rounded-xl shadow-2xl hover:from-cyan-600 hover:to-blue-700 transition-all duration-300 cursor-pointer text-base transform hover:scale-105 hover:shadow-cyan-500/25"
                >
                  <svg className="w-5 h-5 mr-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
                  </svg>
                  Choose Excel File
                </label>
              </div>
            </div>
            
            {/* File Upload Success Message */}
            {selectedFile && (
              <div className="bg-gradient-to-r from-green-500/20 to-emerald-500/20 border border-green-400/30 rounded-xl p-4 sm:p-6 backdrop-blur-sm">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-gradient-to-r from-green-400 to-emerald-500 rounded-full flex items-center justify-center shadow-lg">
                    <span className="text-xl">✅</span>
                  </div>
                  <div>
                    <p className="text-green-100 font-semibold text-base">
                      File uploaded successfully!
                    </p>
                    <p className="text-green-200 text-sm">
                      <span className="font-mono">{selectedFile.name}</span>
                    </p>
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Chart Configuration Section - Only show if file is uploaded */}
          {selectedFile && (
            <>
              {/* Chart Dimension Selection */}
              <div className="space-y-6 relative z-10">
                <h3 className="text-xl sm:text-2xl font-bold bg-gradient-to-r from-purple-400 to-pink-500 bg-clip-text text-transparent flex items-center gap-3">
                  <div className="w-8 h-8 bg-gradient-to-r from-purple-400 to-pink-500 rounded-lg flex items-center justify-center shadow-lg">
                    <span className="text-lg">📐</span>
                  </div>
                  Chart Dimensions
                </h3>
                <div className="flex flex-col sm:flex-row gap-4">
                  {['2d', '3d'].map((dimension) => (
                    <label
                      key={dimension}
                      className={`flex items-center gap-3 p-4 sm:p-6 rounded-xl border-2 cursor-pointer transition-all duration-300 relative overflow-hidden ${
                        chartType === dimension
                          ? 'border-purple-400 bg-gradient-to-r from-purple-400/20 to-pink-400/20 shadow-2xl scale-105'
                          : 'border-purple-300/30 hover:border-purple-400 hover:bg-gradient-to-r hover:from-purple-400/10 hover:to-pink-400/10'
                      }`}
                    >
                      <input
                        type="radio"
                        name="chartType"
                        value={dimension}
                        checked={chartType === dimension}
                        onChange={(e) => {
                          setChartType(e.target.value);
                          setSpecificChart(chartTypes[e.target.value][0].id);
                        }}
                        className="sr-only"
                      />
                      <div className={`w-12 h-12 rounded-lg flex items-center justify-center text-2xl ${
                        chartType === dimension
                          ? 'bg-gradient-to-r from-purple-400 to-pink-500 shadow-lg'
                          : 'bg-purple-400/20'
                      }`}>
                        {dimension === '2d' ? '📊' : '🎯'}
                      </div>
                      <span className="text-cyan-100 font-bold text-lg">{dimension.toUpperCase()}</span>
                    </label>
                  ))}
                </div>
              </div>

              {/* Chart Type Selection */}
              <div className="space-y-6 relative z-10">
                <h3 className="text-xl sm:text-2xl font-bold bg-gradient-to-r from-pink-400 to-orange-500 bg-clip-text text-transparent flex items-center gap-3">
                  <div className="w-8 h-8 bg-gradient-to-r from-pink-400 to-orange-500 rounded-lg flex items-center justify-center shadow-lg">
                    <span className="text-lg">🎨</span>
                  </div>
                  Chart Type
                </h3>
                {chartType === '3d' && (
                  <div className="bg-gradient-to-r from-yellow-500/20 to-orange-500/20 border border-yellow-400/30 rounded-xl p-4 backdrop-blur-sm">
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 bg-gradient-to-r from-yellow-400 to-orange-500 rounded-full flex items-center justify-center">
                        <span className="text-sm">💡</span>
                      </div>
                      <p className="text-yellow-100 text-sm">
                        <strong>Note:</strong> 3D charts are rendered as enhanced 2D versions with depth effects. For true 3D visualization, consider using specialized 3D chart libraries.
                      </p>
                    </div>
                  </div>
                )}
                <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-4">
                  {chartTypes[chartType].map((type) => (
                    <label
                      key={type.id}
                      className={`flex flex-col items-center gap-3 p-4 sm:p-6 rounded-xl border-2 cursor-pointer transition-all duration-300 relative overflow-hidden ${
                        specificChart === type.id
                          ? 'border-pink-400 bg-gradient-to-r from-pink-400/20 to-orange-400/20 shadow-2xl scale-105'
                          : 'border-pink-300/30 hover:border-pink-400 hover:bg-gradient-to-r hover:from-pink-400/10 hover:to-orange-400/10'
                      }`}
                    >
                      <input
                        type="radio"
                        name="specificChart"
                        value={type.id}
                        checked={specificChart === type.id}
                        onChange={(e) => setSpecificChart(e.target.value)}
                        className="sr-only"
                      />
                      <div className={`w-16 h-16 rounded-xl flex items-center justify-center text-3xl ${
                        specificChart === type.id
                          ? 'bg-gradient-to-r from-pink-400 to-orange-500 shadow-lg'
                          : 'bg-pink-400/20'
                      }`}>
                        {type.icon}
                      </div>
                      <span className="text-cyan-100 font-bold text-center text-sm">{type.name}</span>
                    </label>
                  ))}
                </div>
              </div>

              {/* Sheet Selection - Only show if multiple sheets exist */}
              {sheetNames.length > 1 && (
                <div className="space-y-4">
                  <h3 className="text-lg sm:text-xl font-bold text-white flex items-center gap-2">
                    <span>📑</span> Select Sheet
                  </h3>
                  <select
                    value={selectedSheet}
                    onChange={(e) => handleSheetChange(e.target.value)}
                    className="w-full p-3 bg-gray-800 border border-white/20 text-white rounded-lg focus:ring-2 focus:ring-blue-500 text-sm sm:text-base"
                  >
                    {sheetNames.map((sheet, index) => (
                      <option key={index} value={sheet} className="bg-gray-800 text-white">{sheet}</option>
                    ))}
                  </select>
                </div>
              )}

              {/* Data Mapping Section */}
              <div className="space-y-4">
                <h3 className="text-lg sm:text-xl font-bold text-white flex items-center gap-2">
                  <span>🗺️</span> Data Mapping
                </h3>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-300 mb-2">X-Axis (Categories)</label>
                    <select
                      value={xAxis}
                      onChange={(e) => setXAxis(e.target.value)}
                      className="w-full p-3 bg-gray-800 border border-white/20 text-white rounded-lg focus:ring-2 focus:ring-blue-500 text-sm sm:text-base"
                    >
                      <option value="" className="bg-gray-800 text-white">Select X-Axis</option>
                      {columns.map((col, index) => (
                        <option key={index} value={col} className="bg-gray-800 text-white">{col}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-300 mb-2">Y-Axis (Values)</label>
                    <select
                      value={yAxis}
                      onChange={(e) => setYAxis(e.target.value)}
                      className="w-full p-3 bg-gray-800 border border-white/20 text-white rounded-lg focus:ring-2 focus:ring-blue-500 text-sm sm:text-base"
                    >
                      <option value="" className="bg-gray-800 text-white">Select Y-Axis</option>
                      {columns.map((col, index) => (
                        <option key={index} value={col} className="bg-gray-800 text-white">{col}</option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>

              {/* Create Chart Button */}
              <div className="text-center">
                <button
                  onClick={generateChartData}
                  disabled={!xAxis || !yAxis}
                  className="px-8 sm:px-12 py-4 sm:py-6 bg-gradient-to-r from-emerald-500 via-cyan-500 to-blue-600 text-white font-bold rounded-2xl shadow-2xl hover:from-emerald-600 hover:via-cyan-600 hover:to-blue-700 transition-all duration-300 disabled:opacity-50 disabled:cursor-not-allowed text-lg sm:text-xl transform hover:scale-105 hover:shadow-emerald-500/25 relative overflow-hidden group"
                >
                  {/* Animated background */}
                  <div className="absolute inset-0 bg-gradient-to-r from-emerald-400/20 via-cyan-400/20 to-blue-400/20 opacity-0 group-hover:opacity-100 transition-opacity duration-300"></div>
                  
                  <div className="relative z-10 flex items-center justify-center gap-3">
                    <div className="w-8 h-8 bg-white/20 rounded-full flex items-center justify-center">
                      <span className="text-xl">✨</span>
                    </div>
                    <span>Create Amazing Chart</span>
                  </div>
                </button>
              </div>

              {/* Chart Preview */}
              {generatedChart && (
                <div className="mt-6 bg-white/5 border border-white/20 rounded-lg p-4">
                  <h4 className="text-lg font-bold text-white mb-3 flex items-center gap-2">
                    <span>📊</span> Chart Preview ({chartType.toUpperCase()} {specificChart.toUpperCase()})
                  </h4>
                  <div className="w-full h-64 sm:h-80 lg:h-96 relative">
                    {chartType === '3d' ? (
                      // 3D Charts using Three.js
                      <ThreeDChart 
                        data={generatedChart} 
                        chartType={specificChart}
                        width="100%"
                        height="100%"
                        onCanvasReady={(gl, scene, camera) => {
                          threeRendererRef.current = { gl, scene, camera };
                        }}
                      />
                    ) : (
                      // 2D Charts using Chart.js
                      <>
                        {specificChart === 'bar' && (
                          <Bar 
                            data={generatedChart} 
                            options={{ 
                              responsive: true, 
                              maintainAspectRatio: false,
                              plugins: {
                                legend: {
                                  labels: { color: 'white' }
                                }
                              },
                              scales: {
                                x: { ticks: { color: 'white' } },
                                y: { ticks: { color: 'white' } }
                              }
                            }}
                            id="chart-canvas"
                          />
                        )}
                        {specificChart === 'line' && (
                          <Line 
                            data={generatedChart} 
                            options={{ 
                              responsive: true, 
                              maintainAspectRatio: false,
                              plugins: {
                                legend: {
                                  labels: { color: 'white' }
                                }
                              },
                              scales: {
                                x: { ticks: { color: 'white' } },
                                y: { ticks: { color: 'white' } }
                              }
                            }}
                            id="chart-canvas"
                          />
                        )}
                        {specificChart === 'pie' && (
                          <Pie 
                            data={generatedChart} 
                            options={{ 
                              responsive: true, 
                              maintainAspectRatio: false,
                              plugins: {
                                legend: {
                                  labels: { color: 'white' }
                                }
                              }
                            }}
                            id="chart-canvas"
                          />
                        )}
                        {specificChart === 'doughnut' && (
                          <Doughnut 
                            data={generatedChart} 
                            options={{ 
                              responsive: true, 
                              maintainAspectRatio: false,
                              plugins: {
                                legend: {
                                  labels: { color: 'white' }
                                }
                              }
                            }}
                            id="chart-canvas"
                          />
                        )}
                        {specificChart === 'radar' && (
                          <Radar 
                            data={generatedChart} 
                            options={{ 
                              responsive: true, 
                              maintainAspectRatio: false,
                              plugins: {
                                legend: {
                                  labels: { color: 'white' }
                                }
                              },
                              scales: {
                                r: { 
                                  ticks: { color: 'white' },
                                  grid: { color: 'rgba(255,255,255,0.2)' }
                                }
                              }
                            }}
                            id="chart-canvas"
                          />
                        )}
                        {specificChart === 'polar' && (
                          <PolarArea 
                            data={generatedChart} 
                            options={{ 
                              responsive: true, 
                              maintainAspectRatio: false,
                              plugins: {
                                legend: {
                                  labels: { color: 'white' }
                                }
                              },
                              scales: {
                                r: { 
                                  ticks: { color: 'white' },
                                  grid: { color: 'rgba(255,255,255,0.2)' }
                                }
                              }
                            }}
                            id="chart-canvas"
                          />
                        )}
                      </>
                    )}
                  </div>
                  <div className="mt-4 text-center space-y-3">
                    <div className="flex flex-col sm:flex-row gap-3 justify-center">
                      <button
                        onClick={downloadChartAsImage}
                        disabled={isDownloading}
                        className={`px-4 py-2 bg-gradient-to-r from-blue-600 to-purple-700 text-white font-bold rounded-lg transition-all duration-200 text-sm sm:text-base ${
                          isDownloading 
                            ? 'opacity-50 cursor-not-allowed' 
                            : 'hover:from-blue-700 hover:to-purple-800'
                        }`}
                      >
                        <span className="mr-2">
                          {isDownloading ? '⏳' : '🖼️'}
                        </span>
                        {isDownloading ? 'Generating...' : 'Download Image (PNG)'}
                      </button>
                      
                      <button
                        onClick={downloadChartAsPDF}
                        disabled={isDownloading}
                        className={`px-4 py-2 bg-gradient-to-r from-green-600 to-teal-700 text-white font-bold rounded-lg transition-all duration-200 text-sm sm:text-base ${
                          isDownloading 
                            ? 'opacity-50 cursor-not-allowed' 
                            : 'hover:from-green-700 hover:to-teal-800'
                        }`}
                      >
                        <span className="mr-2">
                          {isDownloading ? '⏳' : '📄'}
                        </span>
                        {isDownloading ? 'Generating...' : 'Download PDF'}
                      </button>
                    </div>
                    
                    <p className="text-gray-400 text-xs">
                      💡 Choose Image for high-quality PNG or PDF for document format
                    </p>
                  </div>
                </div>
              )}
            </>
          )}

          {/* Upload History Section */}
          <div className="space-y-4">
            <h2 className="text-xl sm:text-2xl font-bold text-white flex items-center gap-2">
              <span>📋</span> Upload History
            </h2>
            <div className="bg-white/5 rounded-lg p-4">
              {uploadedFiles.length === 0 ? (
                <p className="text-gray-400 text-center py-8 text-sm sm:text-base">No files uploaded yet</p>
              ) : (
                <div className="space-y-3">
                  {uploadedFiles.map((file, index) => (
                    <div
                      key={index}
                      className="flex flex-col sm:flex-row sm:items-center justify-between p-3 bg-white/5 rounded-lg"
                    >
                      <div className="flex items-center gap-3 mb-2 sm:mb-0">
                        <span className="text-xl sm:text-2xl">📄</span>
                        <div>
                          <p className="text-white font-medium text-sm sm:text-base">{file.name}</p>
                          <p className="text-gray-400 text-xs sm:text-sm">
                            {(file.size / 1024).toFixed(1)} KB • {new Date(file.date).toLocaleDateString()}
                          </p>
                        </div>
                      </div>
                      <div className="text-left sm:text-right">
                        <p className="text-gray-300 text-xs sm:text-sm">{file.charts} charts</p>
                        <p className="text-gray-400 text-xs sm:text-sm">{file.downloads} downloads</p>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
    </DashboardLayout>
  );
};

export default UserDashboard; 
