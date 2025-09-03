// Excel Analytics Pro - Professional Data Analysis Platform

class ExcelAnalyticsPlatform {
    constructor() {
        this.currentFile = null;
        this.currentData = null;
        this.currentChart = null;
        this.isAdminLoggedIn = false;
        this.selectedChartType = 'bar';
        
        // Sample data for demonstration
        this.sampleData = [
            {"Month": "January", "Sales": 25000, "Expenses": 18000, "Profit": 7000, "Region": "North"},
            {"Month": "February", "Sales": 32000, "Expenses": 21000, "Profit": 11000, "Region": "South"},
            {"Month": "March", "Sales": 28000, "Expenses": 19500, "Profit": 8500, "Region": "East"},
            {"Month": "April", "Sales": 35000, "Expenses": 23000, "Profit": 12000, "Region": "West"},
            {"Month": "May", "Sales": 42000, "Expenses": 28000, "Profit": 14000, "Region": "North"},
            {"Month": "June", "Sales": 38000, "Expenses": 25000, "Profit": 13000, "Region": "South"}
        ];
        
        this.init();
    }

    init() {
        this.hideLoadingOverlay();
        this.setupEventListeners();
        this.loadStoredData();
        this.updateRecentFiles();
        this.updateAdminStats();
        this.setupMobileOptimizations();
    }

    setupEventListeners() {
        // Navigation
        document.querySelectorAll('.nav-button').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                const page = btn.dataset.page;
                this.showPage(page);
            });
        });

        document.getElementById('backBtn').addEventListener('click', (e) => {
            e.preventDefault();
            this.showPage('home');
        });

        // File upload - Fixed drag and drop + browse functionality
        this.setupFileUpload();

        // Sample data button
        document.getElementById('sampleDataBtn').addEventListener('click', (e) => {
            e.preventDefault();
            this.loadSampleData();
        });

        // Analysis
        document.getElementById('analyzeBtn').addEventListener('click', (e) => {
            e.preventDefault();
            this.startAnalysis();
        });
        
        // Data limit control
        document.getElementById('applyLimitBtn')?.addEventListener('click', (e) => {
            e.preventDefault();
            this.applyDataLimit();
        });

        // Chart type selection
        document.querySelectorAll('.chart-type-button').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                this.selectChartType(btn.dataset.type);
            });
        });

        // Chart configuration
        document.getElementById('xAxisSelect').addEventListener('change', () => this.validateChartConfig());
        document.getElementById('yAxisSelect').addEventListener('change', () => this.validateChartConfig());
        document.getElementById('generateChart').addEventListener('click', (e) => {
            e.preventDefault();
            this.generateChart();
        });

        // Chart export
        document.getElementById('downloadPNG').addEventListener('click', (e) => {
            e.preventDefault();
            this.exportChart('png');
        });
        document.getElementById('downloadPDF').addEventListener('click', (e) => {
            e.preventDefault();
            this.exportChart('pdf');
        });

        // Admin functionality
        document.getElementById('loginBtn').addEventListener('click', (e) => {
            e.preventDefault();
            this.handleAdminLogin();
        });
        document.getElementById('logoutBtn').addEventListener('click', (e) => {
            e.preventDefault();
            this.handleAdminLogout();
        });
        document.getElementById('clearAllBtn').addEventListener('click', (e) => {
            e.preventDefault();
            this.clearAllData();
        });

        // Admin password enter key
        document.getElementById('adminPassword').addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.handleAdminLogin();
            }
        });
    }

    setupFileUpload() {
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');

        if(!dropZone || !fileInput){
            console.error('Upload elements not found in DOM');
            return;
        }

        console.log('[Upload] Setting up file upload functionality');

        // File input change event
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files?.[0];
            if (file) {
                console.log('[Upload] File selected:', file.name, file.type, file.size);
                this.handleFileUpload(file);
            } else {
                console.warn('[Upload] No file selected');
            }
        });

        // Drop zone click to trigger file input
        dropZone.addEventListener('click', (e) => {
            // Only trigger if not clicking directly on the file input
            if (e.target !== fileInput) {
                e.preventDefault();
                fileInput.click();
            }
        });

        // Drag and drop events
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.add('dragover');
            this.animateUploadIcon(true);
        });

        dropZone.addEventListener('dragleave', (e) => {
            e.preventDefault();
            e.stopPropagation();
            // Only remove dragover if we're actually leaving the drop zone
            if (!dropZone.contains(e.relatedTarget)) {
                dropZone.classList.remove('dragover');
                this.animateUploadIcon(false);
            }
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('dragover');
            this.animateUploadIcon(false);
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                this.handleFileUpload(files[0]);
            }
        });

        console.log('[Upload] File upload setup completed successfully');
    }

    setupMobileOptimizations() {
        // Detect if user is on mobile device
        const isMobile = window.innerWidth <= 768 || /Android|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
        
        if (isMobile) {
            // Add mobile class to body for CSS targeting
            document.body.classList.add('mobile-device');
            
            // Optimize chart rendering for mobile
            if (window.Chart) {
                Chart.defaults.responsive = true;
                Chart.defaults.maintainAspectRatio = false;
            }
            
            // Add touch-friendly scrolling for data tables
            const dataPreview = document.getElementById('dataPreview');
            if (dataPreview) {
                dataPreview.style.webkitOverflowScrolling = 'touch';
            }
            
            // Prevent zoom on input focus for iOS
            const inputs = document.querySelectorAll('input[type="text"], input[type="number"], select');
            inputs.forEach(input => {
                if (input.style.fontSize !== '16px') {
                    input.style.fontSize = '16px';
                }
            });
        }
        
        // Handle orientation change
        window.addEventListener('orientationchange', () => {
            setTimeout(() => {
                if (this.currentChart) {
                    this.currentChart.resize();
                }
            }, 100);
        });
        
        // Handle window resize for responsive behavior
        let resizeTimeout;
        window.addEventListener('resize', () => {
            clearTimeout(resizeTimeout);
            resizeTimeout = setTimeout(() => {
                if (this.currentChart) {
                    this.currentChart.resize();
                }
                
                // Update mobile class based on current width
                if (window.innerWidth <= 768) {
                    document.body.classList.add('mobile-device');
                } else {
                    document.body.classList.remove('mobile-device');
                }
            }, 250);
        });
    }    animateUploadIcon(isHover) {
        const icon = document.getElementById('uploadIcon');
        if (isHover) {
            icon.style.transform = 'scale(1.1) rotate(5deg)';
            icon.style.color = '#667eea';
        } else {
            icon.style.transform = 'scale(1) rotate(0deg)';
            icon.style.color = '';
        }
    }

    handleFileUpload(file) {
        // Validate file type
        const validTypes = ['application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
        const validExtensions = ['.xls', '.xlsx'];
        const fileName = file.name.toLowerCase();
        const hasValidExtension = validExtensions.some(ext => fileName.endsWith(ext));
        
        if (!validTypes.includes(file.type) && !hasValidExtension) {
            this.showNotification('Please select a valid Excel file (.xls or .xlsx)', 'error');
            return;
        }

        // Validate file size (10MB limit)
        if (file.size > 10 * 1024 * 1024) {
            this.showNotification('File size too large. Maximum size is 10MB.', 'error');
            return;
        }

        this.currentFile = file;
        this.showUploadProgress();
        
        // Process file with realistic delay
        setTimeout(() => {
            this.processExcelFile(file);
        }, 800);
    }

    showUploadProgress() {
        const progressContainer = document.getElementById('uploadProgress');
        const progressBar = document.getElementById('progressBar');
        const progressPercent = document.getElementById('progressPercent');
        const progressStatus = document.getElementById('progressStatus');
        
        progressContainer.classList.remove('hidden');
        
        const statuses = [
            'Reading file structure...',
            'Parsing Excel data...',
            'Analyzing columns...',
            'Validating data...',
            'Processing complete!'
        ];
        
        let progress = 0;
        let statusIndex = 0;
        
        const interval = setInterval(() => {
            progress += Math.random() * 25;
            if (progress >= 100) {
                progress = 100;
                clearInterval(interval);
            }
            
            progressBar.style.width = progress + '%';
            progressPercent.textContent = Math.round(progress) + '%';
            
            if (statusIndex < statuses.length - 1 && progress > (statusIndex + 1) * 20) {
                statusIndex++;
                progressStatus.textContent = statuses[statusIndex];
            }
        }, 200);
    }

    processExcelFile(file) {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                
                if (jsonData.length === 0) {
                    this.showNotification('The Excel file appears to be empty.', 'error');
                    return;
                }

                this.currentData = jsonData;
                this.showFileDetails(file, jsonData);
                this.saveFileToStorage(file, jsonData);
                this.showNotification('File processed successfully!', 'success');
                
            } catch (error) {
                console.error('Error parsing Excel file:', error);
                this.showNotification('Error parsing Excel file. Please check the file format.', 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    showFileDetails(file, data) {
        const fileDetailsContainer = document.getElementById('fileDetails');
        const fileInfoGrid = document.getElementById('fileInfo');
        
        const columns = Object.keys(data[0] || {});
        const rows = data.length;
        const size = (file.size / 1024).toFixed(1);
        
        fileInfoGrid.innerHTML = `
            <div class="file-info-item">
                <div class="file-info-label">File Name</div>
                <div class="file-info-value">${file.name}</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Size</div>
                <div class="file-info-value">${size} KB</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Rows</div>
                <div class="file-info-value">${rows.toLocaleString()}</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Columns</div>
                <div class="file-info-value">${columns.length}</div>
            </div>
        `;
        
        fileDetailsContainer.classList.remove('hidden');
        document.getElementById('analyzeBtn').disabled = false;
        
        // Hide progress after showing details
        setTimeout(() => {
            document.getElementById('uploadProgress').classList.add('hidden');
        }, 500);
    }

    loadSampleData() {
        this.currentData = this.sampleData;
        this.currentFile = { name: 'Sample_Business_Data.xlsx' };
        
        // Simulate file details for sample data
        const fileDetailsContainer = document.getElementById('fileDetails');
        const fileInfoGrid = document.getElementById('fileInfo');
        
        fileInfoGrid.innerHTML = `
            <div class="file-info-item">
                <div class="file-info-label">File Name</div>
                <div class="file-info-value">Sample_Business_Data.xlsx</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Size</div>
                <div class="file-info-value">2.4 KB</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Rows</div>
                <div class="file-info-value">6</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Columns</div>
                <div class="file-info-value">5</div>
            </div>
        `;
        
        fileDetailsContainer.classList.remove('hidden');
        document.getElementById('analyzeBtn').disabled = false;
        this.showNotification('Sample data loaded successfully!', 'success');
    }

    startAnalysis() {
        if (!this.currentData) {
            this.showNotification('No data available for analysis', 'error');
            return;
        }
        
        // Store the full dataset and use a limited one for analysis if needed
        this.fullDataset = [...this.currentData];
        this.dataLimit = null; // No limit by default
        
        this.populateColumnSelectors();
        this.showDataPreview();
        this.setupDataLimitListeners();
        this.showPage('analysis');
        
        const title = this.currentFile ? `${this.currentFile.name}` : 'Sample Data';
        document.getElementById('analysisTitle').textContent = `Analyzing: ${title}`;
    }
    
    setupDataLimitListeners() {
        const dataLimitInput = document.getElementById('dataLimitInput');
        const applyLimitBtn = document.getElementById('applyLimitBtn');
        
        if (dataLimitInput && applyLimitBtn) {
            dataLimitInput.placeholder = `All (${this.fullDataset.length} rows)`;
            dataLimitInput.max = this.fullDataset.length;
            
            applyLimitBtn.addEventListener('click', () => {
                this.applyDataLimit();
            });
        }
    }
    
    applyDataLimit() {
        const dataLimitInput = document.getElementById('dataLimitInput');
        const limit = parseInt(dataLimitInput.value);
        
        if (isNaN(limit) || limit <= 0) {
            // Reset to all data
            this.dataLimit = null;
            this.currentData = [...this.fullDataset];
            this.showNotification('Analyzing all data rows', 'info');
        } else {
            // Apply limit
            const validLimit = Math.min(limit, this.fullDataset.length);
            this.dataLimit = validLimit;
            this.currentData = this.fullDataset.slice(0, validLimit);
            this.showNotification(`Analyzing first ${validLimit} rows of data`, 'info');
        }
        
        this.showDataPreview();
    }

    populateColumnSelectors() {
        const xAxisSelect = document.getElementById('xAxisSelect');
        const yAxisSelect = document.getElementById('yAxisSelect');
        
        // Clear existing options
        xAxisSelect.innerHTML = '<option value="">Select column...</option>';
        yAxisSelect.innerHTML = '<option value="">Select column...</option>';
        
        if (this.currentData && this.currentData.length > 0) {
            const columns = Object.keys(this.currentData[0]);
            
            columns.forEach(column => {
                const xOption = document.createElement('option');
                xOption.value = column;
                xOption.textContent = column;
                xAxisSelect.appendChild(xOption);
                
                const yOption = document.createElement('option');
                yOption.value = column;
                yOption.textContent = column;
                yAxisSelect.appendChild(yOption);
            });
        }
        
        this.validateChartConfig();
    }

    showDataPreview() {
        const preview = document.getElementById('dataPreview');
        const rowCount = document.getElementById('rowCount');
        
        if (!this.currentData || this.currentData.length === 0) return;

        const columns = Object.keys(this.currentData[0]);
        const previewData = this.currentData.slice(0, 10);
        
        let tableHTML = `
            <table class="data-table">
                <thead>
                    <tr>${columns.map(col => `<th>${col}</th>`).join('')}</tr>
                </thead>
                <tbody>
                    ${previewData.map(row => 
                        `<tr>${columns.map(col => `<td>${row[col] || ''}</td>`).join('')}</tr>`
                    ).join('')}
                </tbody>
            </table>
        `;
        
        preview.innerHTML = tableHTML;
        rowCount.textContent = `${this.currentData.length} rows`;
    }

    selectChartType(type) {
        this.selectedChartType = type;
        
        // Update UI
        document.querySelectorAll('.chart-type-button').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`.chart-type-button[data-type="${type}"]`).classList.add('active');
        
        this.validateChartConfig();
    }

    validateChartConfig() {
        const xAxis = document.getElementById('xAxisSelect').value;
        const yAxis = document.getElementById('yAxisSelect').value;
        const generateBtn = document.getElementById('generateChart');
        
        const canGenerate = xAxis && yAxis && this.selectedChartType;
        generateBtn.disabled = !canGenerate;
        
        if (canGenerate) {
            generateBtn.classList.add('ready');
        } else {
            generateBtn.classList.remove('ready');
        }
    }

    generateChart() {
        const xAxis = document.getElementById('xAxisSelect').value;
        const yAxis = document.getElementById('yAxisSelect').value;
        
        if (!xAxis || !yAxis) {
            this.showNotification('Please select both X and Y axes', 'error');
            return;
        }

        this.showLoadingOverlay();
        
        setTimeout(() => {
            try {
                this.createChart(this.selectedChartType, xAxis, yAxis);
                this.generateInsights(xAxis, yAxis);
                document.getElementById('chartsSection').classList.remove('hidden');
                this.updateStats('charts');
                this.showNotification('Chart generated successfully!', 'success');
                
                // Smooth scroll to chart
                setTimeout(() => {
                    document.getElementById('chartsSection').scrollIntoView({ 
                        behavior: 'smooth', 
                        block: 'start' 
                    });
                }, 300);
            } catch (error) {
                console.error('Chart generation error:', error);
                this.showNotification('Error generating chart. Please try a different chart type or data selection.', 'error');
            } finally {
                this.hideLoadingOverlay();
            }
        }, 800);
    }

    createChart(type, xAxis, yAxis) {
        const canvas = document.getElementById('chartCanvas');
        const container = canvas.parentElement;
        
        // Reset any existing chart or 3D scene
        if (this.currentChart) {
            this.currentChart.destroy();
            this.currentChart = null;
        }
        
        // If we had a 3D chart before, we need to recreate the canvas for Chart.js
        if (canvas.threeDScene) {
            // Remove the 3D canvas and create a new 2D canvas
            canvas.remove();
            const newCanvas = document.createElement('canvas');
            newCanvas.id = 'chartCanvas';
            container.appendChild(newCanvas);
            // Update the canvas reference
            canvas.threeDScene = false;
        }
        
        // Get the updated canvas element (might be a new one)
        const currentCanvas = document.getElementById('chartCanvas');
        const ctx = currentCanvas.getContext('2d');

        const labels = this.currentData.map(row => row[xAxis]);
        const data = this.currentData.map(row => {
            const value = row[yAxis];
            return typeof value === 'number' ? value : parseFloat(value) || 0;
        });

        const colors = ['#1FB8CD', '#FFC185', '#B4413C', '#ECEBD5', '#5D878F', '#DB4545', '#D2BA4C', '#964325', '#944454', '#13343B'];
        
        // Handle 3D chart differently
        if (type === 'bar3d') {
            this.create3DChart(currentCanvas, labels, data, xAxis, yAxis);
            document.getElementById('chartTitle').textContent = `3D ${yAxis} by ${xAxis}`;
            return;
        }
        
        // Determine the Chart.js chart type
        let chartJsType = type;
        
        // Handle scatter charts specially
        if (type === 'scatter') {
            chartJsType = 'scatter';
        }
        
        // Configure dataset based on chart type
        let dataset = {
            label: yAxis,
            data: type === 'scatter' ? 
                data.map((value, index) => ({ x: index, y: value })) : 
                data,
            backgroundColor: type === 'pie' ? 
                colors.slice(0, data.length) : 
                type === 'line' ? 
                    'rgba(31, 184, 205, 0.2)' : 
                    colors[0],
            borderColor: type === 'pie' ? 
                colors.slice(0, data.length) : 
                colors[0],
            borderWidth: type === 'line' ? 3 : 1,
            fill: type === 'line',
            tension: type === 'line' ? 0.4 : 0,
            pointRadius: type === 'scatter' ? 6 : type === 'line' ? 4 : 0,
            pointBackgroundColor: type === 'scatter' || type === 'line' ? colors[0] : undefined,
            pointBorderColor: type === 'scatter' || type === 'line' ? '#ffffff' : undefined,
            pointHoverRadius: type === 'scatter' ? 8 : type === 'line' ? 6 : 0
        };
        
        // For scatter plots, we need special handling
        if (type === 'scatter') {
            dataset = {
                label: yAxis,
                data: labels.map((label, index) => {
                    // For scatter plots, we convert the data to {x, y} format
                    return {
                        x: parseFloat(label) || index,
                        y: data[index]
                    };
                }),
                backgroundColor: colors[0],
                borderColor: colors[0],
                pointRadius: 6,
                pointBackgroundColor: colors[0],
                pointBorderColor: '#ffffff',
                pointHoverRadius: 8
            };
        }
        
        let chartConfig = {
            type: chartJsType,
            data: {
                labels: type === 'scatter' ? null : labels,
                datasets: [dataset]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: {
                    duration: 800,
                    easing: 'easeInOutQuart'
                },
                plugins: {
                    title: {
                        display: true,
                        text: `${yAxis} by ${xAxis}`,
                        font: {
                            size: 16,
                            weight: 'bold'
                        },
                        padding: 20
                    },
                    legend: {
                        display: type === 'pie',
                        position: 'bottom',
                        labels: {
                            padding: 20,
                            usePointStyle: true
                        }
                    },
                    tooltip: {
                        enabled: true,
                        backgroundColor: 'rgba(0, 0, 0, 0.7)',
                        titleFont: {
                            size: 14
                        },
                        bodyFont: {
                            size: 13
                        },
                        padding: 12,
                        displayColors: true,
                        callbacks: {
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed.y !== null) {
                                    label += new Intl.NumberFormat().format(context.parsed.y);
                                }
                                return label;
                            }
                        }
                    }
                },
                scales: type === 'pie' ? undefined : {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: yAxis,
                            font: {
                                weight: 'bold'
                            }
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.05)'
                        },
                        ticks: {
                            callback: function(value) {
                                return new Intl.NumberFormat().format(value);
                            }
                        }
                    },
                    x: {
                        title: {
                            display: true,
                            text: xAxis,
                            font: {
                                weight: 'bold'
                            }
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.05)'
                        }
                    }
                }
            }
        };

        this.currentChart = new Chart(ctx, chartConfig);
        document.getElementById('chartTitle').textContent = `${yAxis} by ${xAxis}`;
    }

    generateInsights(xAxis, yAxis) {
        const insightsContent = document.getElementById('insightsContent');
        
        // Show loading first
        insightsContent.innerHTML = `
            <div class="insights-loading">
                <div class="loading-spinner"></div>
                <p>Generating insights...</p>
            </div>
        `;
        
        setTimeout(() => {
            const insights = this.calculateInsights(xAxis, yAxis);
            insightsContent.innerHTML = insights.map((insight, index) => 
                `<div class="insight-item" style="animation-delay: ${index * 0.1}s">
                    <p>${insight}</p>
                </div>`
            ).join('');
        }, 1500);
    }

    calculateInsights(xAxis, yAxis) {
        const data = this.currentData;
        const values = data.map(row => parseFloat(row[yAxis]) || 0);
        const total = values.reduce((sum, val) => sum + val, 0);
        const average = total / values.length;
        const max = Math.max(...values);
        const min = Math.min(...values);
        const maxIndex = values.indexOf(max);
        const minIndex = values.indexOf(min);
        
        const insights = [
            `📊 The average ${yAxis} across all ${xAxis} categories is ${average.toLocaleString(undefined, {maximumFractionDigits: 2})}.`,
            `🏆 Peak performance was achieved by "${data[maxIndex][xAxis]}" with a ${yAxis} of ${max.toLocaleString()}.`,
            `📉 The lowest ${yAxis} value of ${min.toLocaleString()} was recorded for "${data[minIndex][xAxis]}".`,
            `💰 Total ${yAxis} across all categories amounts to ${total.toLocaleString(undefined, {maximumFractionDigits: 2})}.`,
            `📈 The data shows a range from ${min.toLocaleString()} to ${max.toLocaleString()}, indicating a variance of ${(max - min).toLocaleString()}.`
        ];
        
        return insights;
    }
    
    create3DChart(canvas, labels, data, xAxis, yAxis) {
        // Remove existing canvas content
        const container = canvas.parentElement;
        canvas.remove();
        
        // Create a new canvas for Three.js
        const newCanvas = document.createElement('canvas');
        newCanvas.id = 'chartCanvas';
        container.appendChild(newCanvas);
        
        // Mark this canvas as a 3D scene
        newCanvas.threeDScene = true;
        
        // Initialize Three.js scene
        const scene = new THREE.Scene();
        scene.background = new THREE.Color(0xf8f9fa);
        
        // Set up camera
        const aspectRatio = container.clientWidth / container.clientHeight;
        const camera = new THREE.PerspectiveCamera(60, aspectRatio, 0.1, 1000);
        camera.position.set(20, 20, 20);
        camera.lookAt(0, 0, 0);
        
        // Set up renderer
        const renderer = new THREE.WebGLRenderer({
            canvas: newCanvas,
            antialias: true,
            alpha: true
        });
        renderer.setSize(container.clientWidth, container.clientHeight);
        
        // Add lights
        const ambientLight = new THREE.AmbientLight(0xffffff, 0.5);
        scene.add(ambientLight);
        
        const directionalLight = new THREE.DirectionalLight(0xffffff, 0.8);
        directionalLight.position.set(10, 20, 15);
        scene.add(directionalLight);
        
        // Create a group to hold all bars
        const group = new THREE.Group();
        scene.add(group);
        
        // Find the maximum value to normalize bar heights
        const maxValue = Math.max(...data);
        const barWidth = 1;
        const barDepth = 1;
        const spacing = 0.5;
        const baseColor = new THREE.Color(0x1FB8CD);
        
        // Create bars
        for (let i = 0; i < data.length; i++) {
            const barHeight = (data[i] / maxValue) * 10;
            const geometry = new THREE.BoxGeometry(barWidth, barHeight, barDepth);
            
            // Generate a color with slight variation
            const hue = baseColor.getHSL({}).h;
            const color = new THREE.Color().setHSL(
                (hue + i * 0.05) % 1, 
                0.7, 
                0.5 + (i % 3) * 0.1
            );
            
            const material = new THREE.MeshPhongMaterial({
                color: color,
                flatShading: true,
                transparent: true,
                opacity: 0.9
            });
            
            const bar = new THREE.Mesh(geometry, material);
            
            // Position the bar
            const totalWidth = data.length * (barWidth + spacing) - spacing;
            const startX = -totalWidth / 2;
            bar.position.x = startX + i * (barWidth + spacing) + barWidth / 2;
            bar.position.y = barHeight / 2;
            
            group.add(bar);
            
            // Add label
            if (labels[i]) {
                const textCanvas = document.createElement('canvas');
                textCanvas.width = 100;
                textCanvas.height = 50;
                const textCtx = textCanvas.getContext('2d');
                textCtx.font = '12px Arial';
                textCtx.fillStyle = '#333333';
                textCtx.textAlign = 'center';
                textCtx.fillText(labels[i].toString().substring(0, 10), 50, 20);
                
                const textTexture = new THREE.CanvasTexture(textCanvas);
                const textMaterial = new THREE.MeshBasicMaterial({
                    map: textTexture,
                    transparent: true
                });
                const textGeometry = new THREE.PlaneGeometry(2, 1);
                const textMesh = new THREE.Mesh(textGeometry, textMaterial);
                
                textMesh.position.x = bar.position.x;
                textMesh.position.y = -0.5;
                textMesh.position.z = bar.position.z;
                textMesh.rotation.x = -Math.PI / 2;
                
                group.add(textMesh);
            }
        }
        
        // Add coordinate axes for reference
        const axesHelper = new THREE.AxesHelper(12);
        scene.add(axesHelper);
        
        // Add a grid for better perspective
        const gridHelper = new THREE.GridHelper(20, 20, 0x888888, 0xcccccc);
        scene.add(gridHelper);
        
        // Add orbit controls for interaction
        const controls = new THREE.OrbitControls(camera, renderer.domElement);
        controls.enableDamping = true;
        controls.dampingFactor = 0.1;
        
        // Animation loop
        const animate = () => {
            requestAnimationFrame(animate);
            
            // Very slow rotation for better visualization
            group.rotation.y += 0.0005;
            
            controls.update();
            renderer.render(scene, camera);
        };
        
        animate();
        
        // Handle window resize
        window.addEventListener('resize', () => {
            camera.aspect = container.clientWidth / container.clientHeight;
            camera.updateProjectionMatrix();
            renderer.setSize(container.clientWidth, container.clientHeight);
        });
    }

    exportChart(format) {
        if (!this.currentChart && this.selectedChartType !== 'bar3d') {
            this.showNotification('No chart available to export', 'error');
            return;
        }

        const canvas = document.getElementById('chartCanvas');
        
        if (format === 'png') {
            const link = document.createElement('a');
            link.download = `chart-${Date.now()}.png`;
            link.href = canvas.toDataURL('image/png');
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            this.updateStats('downloads');
            this.showNotification('Chart exported as PNG', 'success');
        } else if (format === 'pdf') {
            const imgData = canvas.toDataURL('image/png');
            const { jsPDF } = window.jspdf;
            const pdf = new jsPDF();
            const imgWidth = 190;
            const imgHeight = (canvas.height * imgWidth) / canvas.width;
            pdf.addImage(imgData, 'PNG', 10, 10, imgWidth, imgHeight);
            pdf.save(`chart-${Date.now()}.pdf`);
        }
        
        this.updateStats('downloads');
        this.showNotification(`Chart exported as ${format.toUpperCase()}`, 'success');
    }

    // Page Management
    showPage(pageName) {
        // Add fade out animation to current page
        const currentPage = document.querySelector('.page.active');
        if (currentPage) {
            currentPage.style.opacity = '0';
            currentPage.style.transform = 'translateY(20px)';
        }
        
        setTimeout(() => {
            // Remove active class from all pages and nav buttons
            document.querySelectorAll('.page').forEach(page => page.classList.remove('active'));
            document.querySelectorAll('.nav-button').forEach(btn => btn.classList.remove('active'));
            
            // Show new page and activate nav button
            const newPage = document.getElementById(`${pageName}Page`);
            const navBtn = document.querySelector(`[data-page="${pageName}"]`);
            
            newPage.classList.add('active');
            navBtn.classList.add('active');
            
            // Reset page styles and add fade in animation
            newPage.style.opacity = '0';
            newPage.style.transform = 'translateY(20px)';
            
            setTimeout(() => {
                newPage.style.opacity = '1';
                newPage.style.transform = 'translateY(0)';
            }, 50);
            
            // Handle admin page specific logic
            if (pageName === 'admin' && !this.isAdminLoggedIn) {
                document.getElementById('adminLogin').classList.remove('hidden');
                document.getElementById('adminDashboard').classList.add('hidden');
            }
        }, 150);
    }

    // Admin Functions
    handleAdminLogin() {
        const password = document.getElementById('adminPassword').value;
        const errorElement = document.getElementById('loginError');
        
        if (password === 'Admin123') {
            this.isAdminLoggedIn = true;
            document.getElementById('adminLogin').classList.add('hidden');
            document.getElementById('adminDashboard').classList.remove('hidden');
            this.updateAdminStats();
            this.updateAdminFileList();
            this.setupAdminEventListeners();
            this.showNotification('Welcome to Admin Dashboard!', 'success');
        } else {
            errorElement.classList.remove('hidden');
            setTimeout(() => errorElement.classList.add('hidden'), 4000);
            this.showNotification('Invalid password. Try "Admin123"', 'error');
        }
    }

    handleAdminLogout() {
        this.isAdminLoggedIn = false;
        document.getElementById('adminLogin').classList.remove('hidden');
        document.getElementById('adminDashboard').classList.add('hidden');
        document.getElementById('adminPassword').value = '';
        this.showNotification('Logged out successfully', 'info');
    }
    
    // Admin file management event listeners
    setupAdminEventListeners() {
        // Refresh files button
        document.getElementById('refreshFilesBtn')?.addEventListener('click', (e) => {
            e.preventDefault();
            this.updateAdminFileList();
            this.showNotification('File list refreshed', 'info');
        });
        
        // Search files input
        document.getElementById('fileSearch')?.addEventListener('input', (e) => {
            this.filterAdminFiles(e.target.value);
        });
        
        // Sort files dropdown
        document.getElementById('fileSortOption')?.addEventListener('change', (e) => {
            this.sortAdminFiles(e.target.value);
        });
    }
    
    filterAdminFiles(searchTerm) {
        const files = this.getStoredFiles();
        const filteredFiles = searchTerm ? 
            files.filter(file => 
                file.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
                file.id.includes(searchTerm)
            ) : files;
        
        this.updateAdminFileListWithData(filteredFiles);
    }
    
    sortAdminFiles(sortOption) {
        const files = this.getStoredFiles();
        let sortedFiles = [...files];
        
        switch (sortOption) {
            case 'date-desc':
                sortedFiles.sort((a, b) => b.uploadDate - a.uploadDate);
                break;
            case 'date-asc':
                sortedFiles.sort((a, b) => a.uploadDate - b.uploadDate);
                break;
            case 'name-asc':
                sortedFiles.sort((a, b) => a.name.localeCompare(b.name));
                break;
            case 'name-desc':
                sortedFiles.sort((a, b) => b.name.localeCompare(a.name));
                break;
            case 'size-desc':
                sortedFiles.sort((a, b) => b.size - a.size);
                break;
            case 'size-asc':
                sortedFiles.sort((a, b) => a.size - b.size);
                break;
        }
        
        this.updateAdminFileListWithData(sortedFiles);
    }

    updateAdminStats() {
        const stats = this.getStorageStats();
        
        // Animate number counting
        this.animateNumber('totalUploads', stats.totalUploads);
        this.animateNumber('totalCharts', stats.totalCharts);
        this.animateNumber('totalDownloads', stats.totalDownloads);
        document.getElementById('storageUsed').textContent = `${stats.storageUsed} MB`;
    }

    animateNumber(elementId, targetValue) {
        const element = document.getElementById(elementId);
        const startValue = 0;
        const duration = 1000;
        const startTime = performance.now();
        
        const animate = (currentTime) => {
            const elapsed = currentTime - startTime;
            const progress = Math.min(elapsed / duration, 1);
            const currentValue = Math.floor(startValue + (targetValue - startValue) * progress);
            
            element.textContent = currentValue;
            
            if (progress < 1) {
                requestAnimationFrame(animate);
            }
        };
        
        requestAnimationFrame(animate);
    }

    updateAdminFileList() {
        const files = this.getStoredFiles();
        this.updateAdminFileListWithData(files);
    }
    
    updateAdminFileListWithData(files) {
        const fileList = document.getElementById('adminFileList');
        
        if (!files || files.length === 0) {
            fileList.innerHTML = '<div class="empty-state"><p>No files to manage</p></div>';
            return;
        }
        
        fileList.innerHTML = files.map(file => `
            <div class="file-item">
                <div class="file-info">
                    <div class="file-info-header">
                        <div class="file-name">${file.name}</div>
                        <div class="file-badge">${file.columns} columns</div>
                    </div>
                    <div class="file-meta">
                        <div class="meta-item">
                            <span class="meta-label">Uploaded:</span>
                            <span class="meta-value">${new Date(file.uploadDate).toLocaleString()}</span>
                        </div>
                        <div class="meta-item">
                            <span class="meta-label">Size:</span>
                            <span class="meta-value">${file.size}</span>
                        </div>
                        <div class="meta-item">
                            <span class="meta-label">Rows:</span>
                            <span class="meta-value">${file.rows.toLocaleString()}</span>
                        </div>
                    </div>
                </div>
                <div class="file-actions">
                    <button class="action-button view-btn" onclick="app.viewFileDetails('${file.id}')">
                        <span class="button-icon">👁️</span>
                        <span>View</span>
                    </button>
                    <button class="action-button analyze-btn" onclick="app.loadFileForAnalysis('${file.id}')">
                        <span class="button-icon">📊</span>
                        <span>Analyze</span>
                    </button>
                    <button class="action-button delete-btn" onclick="app.deleteFile('${file.id}')">
                        <span class="button-icon">🗑️</span>
                        <span>Delete</span>
                    </button>
                </div>
            </div>
        `).join('');
    }
    
    formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' bytes';
        else if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
        else return (bytes / 1048576).toFixed(1) + ' MB';
    }
    
    viewFileDetails(fileId) {
        const files = this.getStoredFiles();
        const file = files.find(f => f.id === fileId);
        
        if (!file) {
            this.showNotification('File not found', 'error');
            return;
        }
        
        // Set modal content
        document.getElementById('modalFileName').textContent = file.name;
        
        // Generate file info
        const modalFileInfo = document.getElementById('modalFileInfo');
        modalFileInfo.innerHTML = `
            <div class="file-info-item">
                <div class="file-info-label">File Name</div>
                <div class="file-info-value">${file.name}</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Upload Date</div>
                <div class="file-info-value">${new Date(file.uploadDate).toLocaleString()}</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Size</div>
                <div class="file-info-value">${file.size}</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Rows</div>
                <div class="file-info-value">${file.rows.toLocaleString()}</div>
            </div>
            <div class="file-info-item">
                <div class="file-info-label">Columns</div>
                <div class="file-info-value">${file.columns}</div>
            </div>
        `;
        
        // Generate data preview
        const dataPreviewWrapper = document.querySelector('#modalDataPreview .data-preview-wrapper');
        
        if (file.data && file.data.length > 0) {
            // Get first 10 rows for preview
            const previewData = file.data.slice(0, 10);
            const headers = Object.keys(previewData[0]);
            
            let tableHTML = '<table class="data-preview-table">';
            
            // Table headers
            tableHTML += '<thead><tr>';
            headers.forEach(header => {
                tableHTML += `<th>${header}</th>`;
            });
            tableHTML += '</tr></thead>';
            
            // Table body
            tableHTML += '<tbody>';
            previewData.forEach(row => {
                tableHTML += '<tr>';
                headers.forEach(header => {
                    tableHTML += `<td>${row[header] !== undefined ? row[header] : ''}</td>`;
                });
                tableHTML += '</tr>';
            });
            tableHTML += '</tbody>';
            tableHTML += '</table>';
            
            dataPreviewWrapper.innerHTML = tableHTML;
        } else {
            dataPreviewWrapper.innerHTML = '<p>No data preview available</p>';
        }
        
        // Set up analyze button
        document.getElementById('modalAnalyzeBtn').onclick = () => {
            this.loadFileForAnalysis(file.id);
            this.closeFileDetailsModal();
        };
        
        // Set up delete button
        document.getElementById('modalDeleteBtn').onclick = () => {
            this.deleteFile(file.id);
            this.closeFileDetailsModal();
        };
        
        // Show the modal
        const modal = document.getElementById('fileDetailsModal');
        modal.classList.add('open');
        
        // Set up close button
        document.getElementById('closeModalBtn').onclick = this.closeFileDetailsModal.bind(this);
        
        // Close modal when clicking outside
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                this.closeFileDetailsModal();
            }
        });
    }
    
    closeFileDetailsModal() {
        const modal = document.getElementById('fileDetailsModal');
        modal.classList.remove('open');
    }
    
    generateFilePreviewTable(data) {
        if (!data || !data.length) return '<p>No data available</p>';
        
        // Limit to first 5 rows for preview
        const previewData = data.slice(0, 5);
        const columns = Object.keys(previewData[0]);
        
        let tableHTML = '<table class="preview-table"><thead><tr>';
        
        // Add headers
        columns.forEach(column => {
            tableHTML += `<th>${column}</th>`;
        });
        
        tableHTML += '</tr></thead><tbody>';
        
        // Add rows
        previewData.forEach(row => {
            tableHTML += '<tr>';
            columns.forEach(column => {
                tableHTML += `<td>${row[column] !== undefined && row[column] !== null ? row[column] : ''}</td>`;
            });
            tableHTML += '</tr>';
        });
        
        tableHTML += '</tbody></table>';
        return tableHTML;
    }

    clearAllData() {
        if (confirm('⚠️ This will permanently delete all uploaded files and statistics. This action cannot be undone.\n\nAre you sure you want to continue?')) {
            localStorage.clear();
            this.updateAdminStats();
            this.updateAdminFileList();
            this.updateRecentFiles();
            this.showNotification('All data cleared successfully', 'success');
        }
    }

    deleteFile(fileId) {
        if (confirm('Are you sure you want to delete this file?')) {
            const files = this.getStoredFiles();
            const updatedFiles = files.filter(file => file.id !== fileId);
            localStorage.setItem('uploadedFiles', JSON.stringify(updatedFiles));
            this.updateAdminFileList();
            this.updateRecentFiles();
            this.updateAdminStats();
            this.showNotification('File deleted successfully', 'success');
        }
    }

    // Storage Functions
    saveFileToStorage(file, data) {
        const fileData = {
            id: Date.now().toString(),
            name: file.name,
            size: (file.size / 1024).toFixed(1) + ' KB',
            uploadDate: new Date().toISOString(),
            rows: data.length,
            columns: Object.keys(data[0] || {}).length,
            data: data
        };

        const existingFiles = this.getStoredFiles();
        existingFiles.push(fileData);
        localStorage.setItem('uploadedFiles', JSON.stringify(existingFiles));
        
        this.updateStats('uploads');
        this.updateRecentFiles();
    }

    getStoredFiles() {
        return JSON.parse(localStorage.getItem('uploadedFiles')) || [];
    }

    getStorageStats() {
        const files = this.getStoredFiles();
        const stats = JSON.parse(localStorage.getItem('platformStats')) || {
            uploads: 0,
            charts: 0,
            downloads: 0
        };

        const storageSize = new Blob([JSON.stringify(files)]).size / (1024 * 1024);
        
        return {
            totalUploads: stats.uploads,
            totalCharts: stats.charts,
            totalDownloads: stats.downloads,
            storageUsed: storageSize.toFixed(2)
        };
    }

    updateStats(type) {
        const stats = JSON.parse(localStorage.getItem('platformStats')) || {
            uploads: 0,
            charts: 0,
            downloads: 0
        };

        stats[type] = (stats[type] || 0) + 1;
        localStorage.setItem('platformStats', JSON.stringify(stats));
        
        if (this.isAdminLoggedIn) {
            this.updateAdminStats();
        }
    }

    loadStoredData() {
        const files = this.getStoredFiles();
        if (files.length > 0) {
            document.getElementById('analyzeBtn').disabled = false;
        }
    }

    updateRecentFiles() {
        const recentFilesContainer = document.getElementById('recentFiles');
        const files = this.getStoredFiles();
        
        if (!files || files.length === 0) {
            recentFilesContainer.innerHTML = `
                <div class="empty-state">
                    <div class="empty-icon">📋</div>
                    <p>No files uploaded yet</p>
                    <button id="sampleDataBtn" class="sample-button">Try Sample Data</button>
                </div>
            `;
            
            // Re-attach event listener for sample data button
            document.getElementById('sampleDataBtn')?.addEventListener('click', () => this.loadSampleData());
            return;
        }
        
        // Clear the container
        recentFilesContainer.innerHTML = '';
        
        // Display the 6 most recent files
        const recentFiles = files.slice(0, 6);
        
        recentFiles.forEach(file => {
            const fileCard = document.createElement('div');
            fileCard.className = 'recent-file-card';
            fileCard.dataset.fileId = file.id;
            
            // Format the date
            const fileDate = new Date(file.date);
            const formattedDate = fileDate.toLocaleDateString();
            
            // Extract file info
            const rowCount = file.data.length;
            const columnCount = file.data[0]?.length || 0;
            
            fileCard.innerHTML = `
                <div class="file-icon">${file.fileType === 'xlsx' ? '📊' : '📑'}</div>
                <div class="recent-file-name" title="${file.name}">${file.name}</div>
                <div class="recent-file-info">
                    <span>${rowCount} rows</span>
                    <span>•</span>
                    <span>${columnCount} columns</span>
                </div>
                <div class="recent-file-date">${formattedDate}</div>
            `;
            
            fileCard.addEventListener('click', () => this.loadFileForAnalysis(file.id));
            
            recentFilesContainer.appendChild(fileCard);
        });
    }

    loadFileForAnalysis(fileId) {
        const files = this.getStoredFiles();
        const file = files.find(f => f.id === fileId);
        
        if (file) {
            this.currentData = file.data;
            this.currentFile = { name: file.name };
            this.startAnalysis();
            this.showNotification(`Loaded ${file.name} for analysis`, 'success');
        }
    }

    // Utility Functions
    showNotification(message, type = 'info') {
        const notification = document.createElement('div');
        notification.className = `notification ${type}`;
        
        const icon = {
            success: '✅',
            error: '❌',
            warning: '⚠️',
            info: 'ℹ️'
        }[type] || 'ℹ️';
        
        notification.innerHTML = `
            <div style="display: flex; align-items: center; gap: 12px;">
                <span style="font-size: 18px;">${icon}</span>
                <span>${message}</span>
            </div>
        `;
        
        document.getElementById('notifications').appendChild(notification);
        
        // Auto remove after 4 seconds
        setTimeout(() => {
            notification.style.opacity = '0';
            notification.style.transform = 'translateX(100%)';
            setTimeout(() => {
                if (notification.parentNode) {
                    notification.parentNode.removeChild(notification);
                }
            }, 300);
        }, 4000);
    }

    showLoadingOverlay() {
        const overlay = document.getElementById('loadingOverlay');
        overlay.classList.remove('hidden');
        overlay.classList.add('show');
    }

    hideLoadingOverlay() {
        const overlay = document.getElementById('loadingOverlay');
        overlay.classList.add('hidden');
        overlay.classList.remove('show');
    }
}

// Initialize the application
let app;

document.addEventListener('DOMContentLoaded', () => {
    app = new ExcelAnalyticsPlatform();
    
    // Add smooth transitions to all page elements
    document.querySelectorAll('.page').forEach(page => {
        page.style.transition = 'opacity 0.3s ease-in-out, transform 0.3s ease-in-out';
    });
    
    // Add interaction feedback to all buttons
    document.querySelectorAll('button:not(.chart-type-btn)').forEach(button => {
        button.addEventListener('mousedown', () => {
            button.style.transform = 'scale(0.95)';
        });
        
        button.addEventListener('mouseup', () => {
            button.style.transform = '';
        });
        
        button.addEventListener('mouseleave', () => {
            button.style.transform = '';
        });
    });
});

// Global helper for inline onclick in HTML
window.openExcelFilePicker = function(){
    if(!app){ console.error('App not initialized yet'); return; }
    const input = document.getElementById('fileInput');
    app.openNativeFileDialog(input);
};

// Initialize modal close button
document.getElementById('closeModalBtn')?.addEventListener('click', () => {
    document.getElementById('fileDetailsModal').classList.remove('open');
});

// Close modal when clicking outside
document.getElementById('fileDetailsModal')?.addEventListener('click', (e) => {
    if (e.target === document.getElementById('fileDetailsModal')) {
        document.getElementById('fileDetailsModal').classList.remove('open');
    }
});

// Handle window resize for responsive charts
window.addEventListener('resize', () => {
    if (app && app.currentChart) {
        app.currentChart.resize();
    }
});