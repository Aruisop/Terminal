import React, { useEffect, useState, useRef } from 'react';
import axios from 'axios';
import { GraphicWalker, GraphicRenderer, PureRenderer } from '@kanaries/graphic-walker';
import * as XLSX from 'xlsx';

const TABS = {
    DESIGN: 'design',
    VIEW: 'view'
};

const ENDPOINTS = {
    DASHBOARD: 'http://localhost:5246/Dashboard',
    DATASET: 'http://localhost:5246/Dataset',
    DATASET_UPLOAD: 'http://localhost:5246/Dataset/Upload'
};

const MESSAGES = {
    NO_DATA: 'Select a dataset from the dropdown or upload from excel',
    LOADING: 'Loading...',
    ERROR: 'An error occurred',
    NO_DASHBOARDS: 'No dashboards saved',
    UPLOAD_SUCCESS: 'File uploaded successfully',
    SAVE_SUCCESS: 'Dashboard saved successfully',
    DATASET_EXISTS: 'Dataset already exists. Please select dataset from dropdown',
    FETCH_ERROR: 'Failed to fetch data',
    EXCEL_ONLY: 'Please upload only Excel files (.xlsx or .xls)',
    DATASET_PATH_EXISTS: 'Excel file with this path already exists. Please select dataset from dropdown',
    UPLOAD_ERROR: 'Failed to upload file',
    INVALID_FILE: 'Invalid file format'
};
const RenderMultipleCharts = ({ singleChartConfig, gwData, index }) => {
    if (!singleChartConfig) {
        console.error('No chart config provided');
        return null;
    }
    const chartConfig = Array.isArray(singleChartConfig) ? singleChartConfig : [singleChartConfig];
    const chartName = singleChartConfig.name || `Chart ${index + 1}`;

    return (
        <div className="chart-item">
            <h3 className="chart-title">{chartName}</h3>
            <GraphicRenderer
                data={gwData.dataSource}
                fields={gwData.fields}
                chart={chartConfig}
            />
        </div>
    );
};

const Chart = () => {
    const [state, setState] = useState({
        dashboard: {
            items: [],
            loading: true,
            error: null,
            selected: null,
            isClicked: false,
            isMultiple: false,
            filteredItems: []
        },
        datasets: {
            items: [],
            loading: true,
            error: null
        },
        gwData: null,
        chartConfig: null,
        activeTab: TABS.DESIGN,
        selectedDataset: '',
    });

    const [isGoClicked, setIsGoClicked] = useState(false);
    const [isGraphicWalkerReady, setIsGraphicWalkerReady] = useState(false);
    const [uploadedFile, setUploadedFile] = useState(null);
    const [isLoading, setIsLoading] = useState(false);
    const gw = useRef(null);

    const fetchData = async (endpoint) => {
        try {
            const response = await axios.get(endpoint);

            switch (endpoint) {
                case ENDPOINTS.DASHBOARD:
                    setState(prev => ({
                        ...prev,
                        dashboard: {
                            ...prev.dashboard,
                            items: response.data,
                            filteredItems: response.data,
                            loading: false,
                            error: null
                        }
                    }));
                    break;

                case ENDPOINTS.DATASET:
                    setState(prev => ({
                        ...prev,
                        datasets: {
                            items: response.data,
                            loading: false,
                            error: null
                        },
                    }));
                    break;

                default:
                    break;
            }

        } catch (err) {
            const errorState = {
                loading: false,
                error: err
            };

            switch (endpoint) {
                case ENDPOINTS.DASHBOARD:
                    setState(prev => ({
                        ...prev,
                        dashboard: { ...prev.dashboard, ...errorState }
                    }));
                    break;

                case ENDPOINTS.DATASET:
                    setState(prev => ({
                        ...prev,
                        datasets: { ...prev.datasets, ...errorState }
                    }));
                    break;
            }

            console.error(`Error fetching ${endpoint}:`, err);
        }
    };

    useEffect(() => {
        fetchData(ENDPOINTS.DATASET);
    }, []);

    useEffect(() => {
        if (state.activeTab === TABS.VIEW) {
            fetchData(ENDPOINTS.DASHBOARD);
        }
    }, [state.activeTab]);

    const handleFileUpload = async (event) => {
        const file = event.target.files[0];
        if (!file) return;
        const fileType = file.name.split('.').pop().toLowerCase();
        if (!['xlsx', 'xls'].includes(fileType)) {
            alert(MESSAGES.EXCEL_ONLY);
            event.target.value = '';
            return;
        }
        const datasetName = file.name.replace(/\.[^/.]+$/, "");
        try {
            const existingDatasets = state.datasets.items;
            const datasetExists = existingDatasets.some(dataset => dataset.datasetName === datasetName);

            if (datasetExists) {
                alert(MESSAGES.DATASET_EXISTS);
                event.target.value = '';
                return;
            }
            const formData = new FormData();
            formData.append('file', file);
            formData.append('datasetName', datasetName);

            await axios.post(ENDPOINTS.DATASET_UPLOAD, formData, {
                headers: {
                    'Content-Type': 'multipart/form-data'
                }
            });
            await fetchData(ENDPOINTS.DATASET);

            setState(prev => ({
                ...prev,
                selectedDataset: datasetName
            }));

            setIsGraphicWalkerReady(false);
            setIsGoClicked(false);
            setState(prev => ({
                ...prev,
                gwData: null
            }));

            const reader = new FileReader();
            reader.onload = (e) => {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const excelData = XLSX.utils.sheet_to_json(firstSheet, {
                    raw: false,
                    defval: null
                });

                if (excelData.length > 0) {
                    const fields = Object.keys(excelData[0]).map(key => ({
                        fid: key,
                        semanticType: !isNaN(excelData[0][key]) ? 'quantitative' : 'nominal',
                        analyticType: 'dimension',
                    }));

                    setTimeout(() => {
                        setState(prev => ({
                            ...prev,
                            gwData: {
                                fields,
                                dataSource: excelData.map(row => {
                                    const cleanRow = {};
                                    Object.keys(row).forEach(key => {
                                        cleanRow[key] = row[key] === null ? '' : row[key];
                                    });
                                    return cleanRow;
                                })
                            }
                        }));
                        setIsGraphicWalkerReady(true);
                        setIsGoClicked(true);
                        setUploadedFile(file);
                    }, 100);
                }
            };

            reader.readAsArrayBuffer(file);
            event.target.value = '';

        } catch (error) {
            if (error.response && error.response.data && error.response.data.errors) {
                const errorMessage = Object.values(error.response.data.errors)
                    .flat()
                    .join('\n');
                alert(errorMessage);
            } else {
                console.error('Error uploading dataset:', error);
                alert('Failed to upload dataset');
            }
            event.target.value = '';
            return;
        }
    };

    const handleGoClick = async () => {
        if (!state.selectedDataset) {
            alert('Please select a dataset');
            return;
        }

        setIsLoading(true);
        setIsGraphicWalkerReady(false);
        const selectedDatasetItem = state.datasets.items.find(dataset => dataset.datasetName === state.selectedDataset);
        if (selectedDatasetItem) {
            if (selectedDatasetItem.isItFromExcel) {
                try {
                    const response = await axios.get(`http://localhost:5246/api/excel/read?excelPath=${encodeURIComponent(selectedDatasetItem.excelPath)}`);
                    const excelData = response.data;

                    if (excelData.length > 0) {
                        const fields = Object.keys(excelData[0]).map(key => ({
                            fid: key,
                            semanticType: typeof excelData[0][key] === 'number' ? 'quantitative' : 'nominal',
                            analyticType: 'dimension',
                        }));

                        setState(prev => ({
                            ...prev,
                            gwData: {
                                fields,
                                dataSource: excelData
                            }
                        }));
                        setIsGraphicWalkerReady(true);
                    } else {
                        setState(prev => ({
                            ...prev,
                            gwData: null
                        }));
                    }
                } catch (error) {
                    console.error('Error fetching Excel data:', error);
                    alert('Failed to fetch data from Excel file.');
                    setState(prev => ({
                        ...prev,
                        gwData: null
                    }));
                }
            } else {
                try {
                    const response = await axios.get(`http://localhost:5246/api/storedprocedure/execute?storedProcedureName=${encodeURIComponent(selectedDatasetItem.sp)}`);
                    const spData = response.data;

                    if (spData.length > 0) {
                        const fields = Object.keys(spData[0]).map(key => ({
                            fid: key,
                            semanticType: typeof spData[0][key] === 'number' ? 'quantitative' : 'nominal',
                            analyticType: 'dimension',
                        }));

                        setState(prev => ({
                            ...prev,
                            gwData: {
                                fields,
                                dataSource: spData
                            }
                        }));
                        setIsGraphicWalkerReady(true);
                    } else {
                        setState(prev => ({
                            ...prev,
                            gwData: null
                        }));
                    }
                } catch (error) {
                    console.error('Error fetching stored procedure data:', error);
                    alert('Failed to fetch data from stored procedure.');
                    setState(prev => ({
                        ...prev,
                        gwData: null
                    }));
                }
            }

            setIsGoClicked(true);
        } else {
            alert('Selected dataset not found.');
            setIsGoClicked(false);
        }
        setIsLoading(false);
    };

    const exportChartConfigs = async () => {
        if (!gw.current) return;

        const dashboardName = prompt("Enter the name for the dashboard:");
        if (!dashboardName) return;

        try {
            const chartConfigList = gw.current.exportCode();
            const isMultiple = Array.isArray(chartConfigList) && chartConfigList.length > 1;
            const datasetNameToUse = state.selectedDataset ||
                (uploadedFile ? state.datasets.items.find(d => d.isItFromExcel)?.datasetName : '');

            if (!datasetNameToUse) {
                alert('No dataset selected');
                return;
            }

            await axios.post(ENDPOINTS.DASHBOARD, {
                dashboardName: dashboardName,
                datasetName: datasetNameToUse,
                jsonFormat: JSON.stringify(chartConfigList, null, 2),
                isMultiple,
            });

            alert(MESSAGES.SAVE_SUCCESS);
            if (state.activeTab === TABS.VIEW) {
                await fetchData(ENDPOINTS.DASHBOARD);
            }
            setUploadedFile(null);

        } catch (error) {
            console.error('Error saving dashboard:', error);
            alert(error.response?.data?.error || 'Failed to save dashboard');
        }
    };


    const handleDashboardAction = async (dashboard, action) => {
        if (action === 'view') {
            setState(prev => ({
                ...prev,
                dashboard: {
                    ...prev.dashboard,
                    loading: true
                }
            }));

            const selectedDatasetItem = state.datasets.items.find(
                dataset => dataset.datasetName === dashboard.datasetName
            );

            if (selectedDatasetItem) {
                try {
                    let responseData;
                    if (selectedDatasetItem.isItFromExcel) {
                        const response = await axios.get(`http://localhost:5246/api/excel/read?excelPath=${encodeURIComponent(selectedDatasetItem.excelPath)}`);
                        responseData = response.data;
                    } else {
                        const response = await axios.get(`http://localhost:5246/api/storedprocedure/execute?storedProcedureName=${encodeURIComponent(selectedDatasetItem.sp)}`);
                        responseData = response.data;
                    }

                    if (responseData.length > 0) {
                        const fields = Object.keys(responseData[0]).map(key => ({
                            fid: key,
                            semanticType: typeof responseData[0][key] === 'number' ? 'quantitative' : 'nominal',
                            analyticType: 'dimension',
                        }));

                        setState(prev => ({
                            ...prev,
                            dashboard: {
                                ...prev.dashboard,
                                selected: dashboard,
                                isClicked: true,
                                isMultiple: dashboard.isMultiple,
                                loading: false
                            },
                            chartConfig: JSON.parse(dashboard.jsonFormat),
                            gwData: {
                                fields,
                                dataSource: responseData
                            }
                        }));
                    } else {
                        setState(prev => ({
                            ...prev,
                            dashboard: {
                                ...prev.dashboard,
                                loading: false
                            }
                        }));
                        alert('No data available for this dashboard');
                    }
                } catch (error) {
                    console.error('Error fetching dataset data:', error);
                    setState(prev => ({
                        ...prev,
                        dashboard: {
                            ...prev.dashboard,
                            loading: false
                        }
                    }));
                    alert('Failed to fetch dataset data');
                }
            } else {
                setState(prev => ({
                    ...prev,
                    dashboard: {
                        ...prev.dashboard,
                        loading: false
                    }
                }));
                alert('Dataset not found');
            }
        } else {
            setState(prev => ({
                ...prev,
                dashboard: {
                    ...prev.dashboard,
                    selected: dashboard
                }
            }));
        }
    };


    const resetDashboardView = () => {
        setState(prev => ({
            ...prev,
            dashboard: {
                ...prev.dashboard,
                isClicked: false,
                isMultiple: false,
                selected: null
            }
        }));
    };

    const handleDatasetChange = (e) => {
        setState(prev => ({
            ...prev,
            selectedDataset: e.target.value
        }));
        setIsGoClicked(false);
        setIsGraphicWalkerReady(false);
    };

    const handleGraphicWalkerLoad = () => {
        setIsGraphicWalkerReady(true);
    };

    const curLang = 'english';
    if (state.datasets.loading) {
        return (
            <div className='loading-overlay'>
                <div className="loading-spinner"></div>
                <div className="loading-text">{MESSAGES.LOADING}</div>
            </div>
        );
    }

    if (state.datasets.error) {
        return (
            <div className="error-message">
                {MESSAGES.ERROR}: {state.datasets.error.message}
            </div>
        );
    }
    return (
        <div className="chart-page">
            <div className="tabs">
                {Object.values(TABS).map(tab => (
                    <button
                        key={tab}
                        className={`tab-button ${state.activeTab === tab ? 'active' : ''}`}
                        onClick={() => setState(prev => ({ ...prev, activeTab: tab }))}
                    >
                        {tab.charAt(0).toUpperCase() + tab.slice(1)} Dashboards
                    </button>
                ))}
            </div>

            <div className="content">
                {state.activeTab === TABS.DESIGN && (
                    <div className="design-tab">
                        <div style={{ marginBottom: '10px', display: 'flex', alignItems: 'center', gap: '10px' }}>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flex: 1 }}>
                                <label htmlFor="dataset-select" style={{ fontWeight: 'bold' }}>Dataset Name:</label>
                                <select
                                    id="dataset-select"
                                    className="dataset-dropdown"
                                    value={state.selectedDataset}
                                    onChange={handleDatasetChange}
                                    disabled={isLoading}
                                >
                                    <option value="">Select</option>
                                    {state.datasets.items.map((dataset, idx) => (
                                        <option key={dataset.id || idx} value={dataset.datasetName}>
                                            {dataset.datasetName}
                                        </option>
                                    ))}
                                </select>
                                <button
                                    className="go-button"
                                    onClick={handleGoClick}
                                    disabled={!state.selectedDataset || isLoading}
                                >
                                    {isLoading ? MESSAGES.LOADING : 'Go'}
                                </button>
                            </div>
                            <div className="upload-section">
                                <input
                                    type="file"
                                    accept=".xlsx,.xls"
                                    onChange={handleFileUpload}
                                    style={{ display: 'none' }}
                                    id="file-upload"
                                    disabled={isLoading}
                                />
                                <label
                                    htmlFor="file-upload"
                                    className={`upload-button ${isLoading ? 'disabled' : ''}`}
                                >
                                    Upload Excel
                                </label>
                            </div>
                        </div>

                        {isGraphicWalkerReady && (
                            <button className="export-button" onClick={exportChartConfigs}>
                                Export Dashboard
                            </button>
                        )}
                        <div className="chart-container">
                            {isGoClicked && state.gwData ? (
                                <GraphicWalker
                                    fields={state.gwData.fields}
                                    data={state.gwData.dataSource}
                                    vizThemeConfig="default"
                                    storeRef={gw}
                                    i18nLang={curLang}
                                    i18nResources={MESSAGES}
                                    appearance='media'
                                    onLoad={handleGraphicWalkerLoad}
                                />
                            ) : (
                                <div className="no-data-message">
                                    {MESSAGES.NO_DATA}
                                </div>
                            )}
                            {state.datasets.loading && (
                                <div className="loading-overlay">
                                    <div className="loading-spinner"></div>
                                    <div className="loading-text">{MESSAGES.LOADING}</div>
                                </div>
                            )}
                        </div>
                    </div>
                )}
                {state.activeTab === TABS.VIEW && (
                    <div className="view-tab">
                        {!state.dashboard.isClicked ? (
                            <div className="dashboard-list">
                                <h3>Saved Dashboards</h3>
                                {state.dashboard.items.length > 0 ? (
                                    <>
                                        <ul className="dashboard-items">
                                            {state.dashboard.items.map(dashboard => (
                                                <li key={dashboard.id} className="dashboard-item">
                                                    <input
                                                        type="radio"
                                                        name="selectedDashboard"
                                                        value={dashboard.id}
                                                        onChange={() => handleDashboardAction(dashboard, 'select')}
                                                        id={`dashboard-${dashboard.id}`}
                                                    />
                                                    <label htmlFor={`dashboard-${dashboard.id}`}>
                                                        {dashboard.dashboardName}
                                                    </label>
                                                </li>
                                            ))}
                                        </ul>
                                        <button
                                            className="view-button"
                                            onClick={() => state.dashboard.selected && handleDashboardAction(state.dashboard.selected, 'view')}
                                            disabled={!state.dashboard.selected}
                                        >
                                            {state.dashboard.loading ? MESSAGES.LOADING : 'Show Dashboard'}
                                        </button>
                                    </>
                                ) : (
                                    <div className="no-dashboards-message">
                                        {MESSAGES.NO_DASHBOARDS}
                                    </div>
                                )}
                            </div>
                        ) : (
                            state.chartConfig && state.gwData && (
                                <div className="chart-container">
                                    {state.dashboard.loading ? (
                                        <div className="loading-overlay">
                                            <div className="loading-spinner"></div>
                                            <div className="loading-text">{MESSAGES.LOADING}</div>
                                        </div>
                                    ) : state.dashboard.isMultiple ? (
                                        <div className="multiple-charts">
                                            {(() => {
                                                const charts = [];
                                                for (let i = 0; i < state.chartConfig.length; i++) {
                                                    charts.push(
                                                        <RenderMultipleCharts
                                                            key={i}
                                                            singleChartConfig={state.chartConfig[i]}
                                                            gwData={state.gwData}
                                                            index={i}
                                                        />
                                                    );
                                                }
                                                return charts;
                                            })()}
                                        </div>
                                    ) : (
                                        <GraphicRenderer
                                            data={state.gwData.dataSource}
                                            fields={state.gwData.fields}
                                            chart={state.chartConfig}
                                        />
                                    )}
                                    <button className="back-button" onClick={resetDashboardView}>
                                        Back to View Dashboard
                                    </button>
                                </div>
                            )
                        )}
                    </div>
                )}
            </div>
        </div>
    );
};

export default Chart;
