import * as React from "react";
import { useState, useEffect } from "react";
import { useHistory, Link } from "react-router-dom";
import { ISonaDeviationProps } from "../ISonaDeviationProps";
import "./Css/InitiatorDashboard.scss"
import SPCRUDOPS from "../../service/BAL/spcrud";
import * as XLSX from 'xlsx';

import Left from "../../assets/LeftArrow.png";
import Right from "../../assets/RightArrow.png";

import { sp } from "@pnp/sp/presets/all";

import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import Edit from "../../assets/Pencil.png";
import View from "../../assets/Eye.png";
import { Dropdown } from "@fluentui/react";
export const ReportDashboard: React.FC<ISonaDeviationProps> = (props: ISonaDeviationProps) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [listData, setListData] = useState<any[]>([]);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [loading, setLoading] = useState(true);
    const [filters, setFilters] = useState({
        PartNo: "",
        DeviationNo: "",
        SupplierName: "",
        BatchNo: "",
        PlantName: "",
        RequestorName: "",
        NextApproverName: "",
        DeviationType: "",
        Status: ""
    });

    // Pagination
    const itemsPerPage = 10;
    const [currentPage, setCurrentPage] = useState(1);
    const totalPages = Math.ceil(filteredData.length / itemsPerPage);

    const history = useHistory();

    const formatDate = (dateString: string) => {
        if (!dateString) return "";
        const date = new Date(dateString);

        return date
            .toLocaleDateString("en-GB", {
                day: "2-digit",
                month: "2-digit",
                year: "numeric",
            })
            .replace(/\//g, "-"); // dd-MM-yyyy
    };

    const getUniqueOptions = (data: any[], field: string) => {
        return Array.from(
            new Set(data.map(item => item[field]).filter(Boolean))
        ).map(value => ({
            key: value,
            text: value
        }));
    };
    const partNoOptions = getUniqueOptions(listData, "PartNo");
    const deviationNoOptions = getUniqueOptions(listData, "DeviationNo");
    const supplierOptions = getUniqueOptions(listData, "SupplierName");
    const plantOptions = getUniqueOptions(listData, "PlantName");
    const deviationTypeOptions = getUniqueOptions(listData, "DeviationType");
    const statusOptions = getUniqueOptions(listData, "Status");

    React.useEffect(() => {
        sp.setup({
            spfxContext: props.currentSPContext
        });
    }, []);

    //Latest 
    const GetListData = async () => {

        setLoading(true);

        const spCrudOps = await SPCRUDOPS();

        // 🔹 Get logged-in user
        const loggedInUser = await sp.web.currentUser();

        const parentItems = await spCrudOps.getData(
            "DeviationDetails",
            "*,ID,PartNo,DeviationNo,SupplierName,BatchNo,PlantName,ProductionQTY,RequestorName,Author/ID,NextApprover/ID,NextApprover/Title,Author/Title",
            "Author,NextApprover",
            "",
            { column: "ID", isAscending: true },
            5000,
            props
        );
        const data = parentItems
        setListData(data);
        setFilteredData(data);
        setLoading(false);
    };

    useEffect(() => {
        GetListData();
    }, []);

    // 🔹 Filter data by search term
    useEffect(() => {
        const lowerSearch = searchTerm.toLowerCase();
        const filtered = listData.filter(item => {
            // 🔹 GLOBAL SEARCH (across all fields)
            const globalMatch = !searchTerm || Object.values({
                PartNo: item.PartNo,
                DeviationNo: item.DeviationNo,
                SupplierName: item.SupplierName,
                BatchNo: item.BatchNo,
                PlantName: item.PlantName,
                RequestorName: item.RequestorName,
                NextApproverName: item.NextApprover?.Title,
                DeviationType: item.DeviationType,
                Status: item.Status
            })
                .join(" ")
                .toLowerCase()
                .includes(lowerSearch);

            // 🔹 COLUMN FILTERS (individual fields)
            const columnMatch =
                (!filters.PartNo || item.PartNo?.toLowerCase().includes(filters.PartNo.toLowerCase())) &&
                (!filters.DeviationNo || item.DeviationNo?.toLowerCase().includes(filters.DeviationNo.toLowerCase())) &&
                (!filters.SupplierName || item.SupplierName?.toLowerCase().includes(filters.SupplierName.toLowerCase())) &&
                (!filters.BatchNo || item.BatchNo?.toLowerCase().includes(filters.BatchNo.toLowerCase())) &&
                (!filters.PlantName || item.PlantName?.toLowerCase().includes(filters.PlantName.toLowerCase())) &&
                (!filters.RequestorName || item.RequestorName?.toLowerCase().includes(filters.RequestorName.toLowerCase())) &&
                (!filters.NextApproverName || item.NextApprover?.Title?.toLowerCase().includes(filters.NextApproverName.toLowerCase())) &&
                (!filters.DeviationType || item.DeviationType?.toLowerCase().includes(filters.DeviationType.toLowerCase())) &&
                (!filters.Status || item.Status?.toLowerCase().includes(filters.Status.toLowerCase()));

            return globalMatch && columnMatch;
        });
        setFilteredData(filtered);
        setCurrentPage(1);
    }, [searchTerm, filters, listData]);

    const handleFilterChange = (field: string, value: string) => {
        setFilters(prev => ({
            ...prev,
            [field]: value
        }));
    };

    const handleReset = () => {
        setSearchTerm("");
        setFilters({
            PartNo: "",
            DeviationNo: "",
            SupplierName: "",
            BatchNo: "",
            PlantName: "",
            RequestorName: "",
            NextApproverName: "",
            DeviationType: "",
            Status: ""
        });
    };
    const addAllOption = (options: any[]) => [
        { key: "SELECT", text: "Select" },
        ...options
    ];

    // 🔹 Pagination
    const handlePageChange = (page: number) => {
        if (page >= 1 && page <= totalPages) {
            setCurrentPage(page);
        }
    };

    const exportToExcel = () => {
        // Always export all filtered data (ignore pagination)
        const dataToExport = [...filteredData].sort((a, b) => b.ID - a.ID);

        if (dataToExport.length === 0) {
            alert("No records found to export.");
            return;
        }

        // Map fields to clean column labels
        const exportData = dataToExport.map((item) => ({
            "Part No": item.PartNo,
            "Deviation On": item.DeviationNo,
            "Supplier Name": item.SupplierName,
            "Batch No": item.BatchNo,
            "Plant Name": item.PlantName,
            "Requestor Name": item.RequestorName,
            "Next Approver": item.NextApprover?.Title,
            "Deviation Type": item.DeviationType,
            "Status": item.Status
        }));

        // Create sheet + workbook
        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Report Dashboard");

        // Save file with today’s date
        const today = new Date().toISOString().slice(0, 10); // yyyy-mm-dd
        XLSX.writeFile(workbook, `ReportDashboard_${today}.xlsx`);
    };

    const sortedData = [...filteredData].sort((a, b) => b.ID - a.ID);

    const paginatedData = sortedData.slice(
        (currentPage - 1) * itemsPerPage,
        currentPage * itemsPerPage
    );

    return (
        <div className="min-h-screen bg-gray-100">
            {/* Header */}
            <div className="header">
                <div className="left-banner">
                    <div className="logo-text">
                        <h2>Report Dashboard</h2>
                    </div>
                </div>
            </div>
            <div className='col-md-12 px-2 d-flex justify-content-between align-items-center flex-wrap' style={{ margin: ".3rem" }}>
                <div>
                    <input type="text" placeholder="Search..."
                        className="form-control" style={{ width: "250px" }}
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                    />
                </div>
            </div>
            <div className="row mb-2 mt-2 ml-auto" style={{ marginTop: '0.5rem', marginLeft: 'auto', marginRight: 'auto'}}>
                <div className="col">
                    <Dropdown
                        placeholder="Part No"
                        options={partNoOptions}
                        selectedKey={filters.PartNo || "SELECT"}
                        onChange={(e, option) =>
                            handleFilterChange("PartNo", option?.key as string)
                        }
                    />
                </div>

                <div className="col">
                    <Dropdown
                        placeholder="Deviation No"
                        options={deviationNoOptions}
                        selectedKey={filters.DeviationNo || "SELECT"}
                        onChange={(e, option) =>
                            handleFilterChange("DeviationNo", option?.key as string)
                        }
                    />
                </div>

                {/* <div className="col">
                    <Dropdown
                        placeholder="Department Name"
                        options={supplierOptions}
                        selectedKey={filters.SupplierName || "SELECT"}
                        onChange={(e, option) =>
                            handleFilterChange("SupplierName", option?.key as string)
                        }
                    />
                </div> */}

                {/* <div className="col">
                    <input placeholder="Batch No" className="form-control"
                        value={filters.BatchNo}
                        onChange={(e) => handleFilterChange("BatchNo", e.target.value)}
                    />
                </div> */}

                <div className="col">
                    <Dropdown
                        placeholder="Plant Name"
                        options={plantOptions}
                        selectedKey={filters.PlantName || "SELECT"}
                        onChange={(e, option) =>
                            handleFilterChange("PlantName", option?.key as string)
                        }
                    />
                </div>

                {/* <div className="col">
                    <input placeholder="Requestor Name" className="form-control"
                        value={filters.RequestorName}
                        onChange={(e) => handleFilterChange("RequestorName", e.target.value)}
                    />
                </div> */}
            </div>

            <div className="row mb-2 mt-2 ml-auto" style={{ marginTop: '0.5rem', marginLeft: 'auto', marginRight: 'auto'}}>

                <div className="col">
                    <Dropdown
                        placeholder="Deviation Type"
                        options={deviationTypeOptions}
                        selectedKey={filters.DeviationType || "SELECT"}
                        onChange={(e, option) =>
                            handleFilterChange("DeviationType", option?.key as string)
                        }
                    />
                </div>

                <div className="col">
                    <Dropdown
                        placeholder="Status"
                        options={statusOptions}
                        selectedKey={filters.Status || "SELECT"}
                        onChange={(e, option) =>
                            handleFilterChange("Status", option?.key as string)
                        }
                    />
                </div>

                <div className="col">
                    <button
                        className="btn btn-secondary"
                        onClick={handleReset}
                    >
                        Reset
                    </button>
                </div>
                {filteredData.length > 0 && (
                    <div className="col">
                        <button
                            className="btn btn-secondary"
                            onClick={exportToExcel}
                        >
                            Export 
                        </button>
                    </div>
                )}
            </div>
            <main className="Main-Dash mx-2">
                <div className="overflow-x-auto">
                    <div className="table-vert-scroll">

                        <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                            <thead
                                style={{ backgroundColor: "#3c3e45" }}
                                className="text-white"
                            >
                                <tr>
                                    <th className="px-4 py-2">Part No.</th>
                                    <th className="px-4 py-2">Deviation No</th>
                                    {/* <th className="px-4 py-2">Supplier Name</th> */}
                                    <th className="px-4 py-2">Batch / Lot No </th>
                                    <th className="px-4 py-2">Plant Name</th>
                                    <th className="px-4 py-2">Requester Name</th>
                                    <th className="px-4 py-2">Next Approver</th>
                                    <th className="px-4 py-2">Deviation Type</th>
                                    <th className="px-4 py-2">Status</th>
                                    <th className="px-4 py-2">View</th>
                                </tr>
                            </thead>
                            <tbody>
                                {paginatedData.map((item, index) => (
                                    <tr key={index} className="border-t">
                                        <td className="px-4 py-2">{item.PartNo}</td>
                                        <td className="px-4 py-2">{item.DeviationNo}</td>
                                        {/* <td className="px-4 py-2">{item.SupplierName}</td> */}
                                        <td className="px-4 py-2">{item.BatchNo}</td>
                                        <td className="px-4 py-2">{item.PlantName}</td>
                                        <td className="px-4 py-2">{item.RequestorName}</td>
                                        <td className="px-4 py-2">{item.NextApprover?.Title}</td>
                                        <td className="px-4 py-2">{item.DeviationType}</td>
                                        <td className="px-4 py-2">{item.Status}</td>
                                        <td className="px-4 py-2">
                                            {item.Status === "Send Back" || item.Status === "Saved as Draft" ? (
                                                <Link to={`/DeviationEditForm/${item.Id}`}>
                                                    <img src={Edit} width={15} alt="Edit" />
                                                </Link>
                                            ) : (
                                                <>
                                                    <Link to={`/DeviationViewForm/${item.Id}`}>
                                                        <img src={View} width={15} alt="View" />
                                                    </Link>
                                                </>
                                            )}
                                        </td>
                                        {/* <td className="px-4 py-2">

                                            {item.Status !== "Pending Amendment" && (
                                                <>
                                                    <Link to={`/BGApproverForm/${item.Id}`}>
                                                        <img src={Edit} width={15} />
                                                    </Link>
                                                    &nbsp;&nbsp;
                                                    <Link to={`/BGViewmore/${item.Id}`}>
                                                        <img src={View} width={15} />
                                                    </Link>
                                                </>
                                            )}
                                            {item.Status === "Pending Amendment" && (
                                                <>
                                                    <Link to={`/BGAmmendmentApproverForm/${item.Id}`}>
                                                        <img src={Edit} width={15} />
                                                    </Link>
                                                    &nbsp;&nbsp;
                                                    <Link to={`/BGAmmendmentViewmore/${item.Id}`}>
                                                        <img src={View} width={15} />
                                                    </Link>
                                                </>
                                            )}

                                        </td> */}
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    {/* Pagination */}
                    <div className="flex justify-center mt-6 overflow-x-auto">
                        <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-#2149d5 rounded shadow" style={{ textAlign: "end" }}>
                            {/* Previous Button */}
                            <button
                                onClick={() => handlePageChange(currentPage - 1)}
                                disabled={currentPage === 1}
                                style={{
                                    backgroundColor: "#fff",
                                    border: "1px solid #000 !important",
                                    marginRight: "5px",
                                    opacity: currentPage === 1 ? 0.5 : 1,
                                }}
                                className="px-3 py-1 border rounded"
                            >
                                <img src={Left} alt="" width={15} />
                            </button>
                            {/* Main Page Numbers */}
                            {Array.from({ length: totalPages }, (_, i) => i + 1)
                                .filter((page) => Math.abs(page - currentPage) <= 2)
                                .map((page) => (
                                    <button
                                        key={page}
                                        onClick={() => handlePageChange(page)}
                                        style={{
                                            backgroundColor: currentPage === page ? "#3c3e45" : "#fff",
                                            color: currentPage === page ? "#fff" : "#000",
                                            fontWeight: currentPage === page ? "bold" : "normal",
                                            margin: currentPage === page ? "5px" : "5px",
                                        }}
                                        className="px-3 py-1 border rounded"
                                    >
                                        {page}
                                    </button>
                                ))}

                            {/* Next Button */}
                            <button
                                onClick={() => handlePageChange(currentPage + 1)}
                                disabled={currentPage === totalPages}
                                style={{
                                    backgroundColor: "#fff",
                                    border: "1px solid #000 !important",
                                    marginLeft: "5px",
                                    opacity: currentPage === totalPages ? 0.5 : 1,
                                }}
                                className="px-3 py-1 border rounded"
                            >
                                <img src={Right} alt="" width={15} />
                            </button>
                        </div>
                    </div>
                </div>
            </main>
        </div>
    );
};

export default ReportDashboard;
