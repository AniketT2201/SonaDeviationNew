import * as React from "react";
import { useState, useEffect } from "react";
import { useHistory, Link } from "react-router-dom";
import type { ISonaDeviationProps } from '../ISonaDeviationProps';
import "./Css/InitiatorDashboard.scss";
import SPCRUDOPS from "../../service/BAL/spcrud";

import { sp } from "@pnp/sp/presets/all";

import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import Left from "../../assets/LeftArrow.png";
import Right from "../../assets/RightArrow.png";
 
import View from "../../assets/Eye.png";
import Edit from "../../assets/Pencil.png";
import Renew from "../../assets/Renew.png";
export const COODashboard: React.FC<ISonaDeviationProps> = (props: ISonaDeviationProps) => {
    const [searchTerm, setSearchTerm] = useState("");
    const [listData, setListData] = useState<any[]>([]);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [loading, setLoading] = useState(true);

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

    //Latest 
    const GetListData = async () => {

        setLoading(true);

        const spCrudOps = await SPCRUDOPS();

        // 🔹 Get logged-in user
        const loggedInUser = await sp.web.currentUser();

        const parentItems = await spCrudOps.getRootData(
            "DeviationDetails",
            "ID,PartNo,DeviationNo,SupplierName,BatchNo,PlantName,ProductionQTY,RequestorName,ProductionDate,Status,Stage,DeviationType,NextApprover/ID,ApprovalMatrix",
            "NextApprover",
            "",
            { column: "ID", isAscending: true },
            5000,
            props
        );
        const data = parentItems.filter((item: any) => {
            if (item.NextApprover?.ID !== loggedInUser.Id) return false;

            let matrix: any[] = [];

            try {
                matrix = JSON.parse(item.ApprovalMatrix || "[]");
            } catch (e) {
                console.error("Invalid ApprovalMatrix JSON", item);
                return false;
            }

            const engineeringStep = matrix.find(
                (m: any) => m.Role === "COO" && m.required === true
            );

            return engineeringStep?.Stage === item.Stage;
        });

        setListData(data);
        setFilteredData(data);
        setLoading(false);
    };

    useEffect(() => {
        GetListData();
    }, []);

    // 🔹 Filter data by search term

    useEffect(() => {
        if (!searchTerm) {
            setFilteredData(listData);
            setCurrentPage(1);
            return;
        }

        const lowerSearch = searchTerm.toLowerCase();

        const filtered = listData.filter((item) =>
            Object.values({
                BGReqtNo: item.BGReqtNo,
                BeneficiaryName: item.BeneficiaryName,
                BGType: item.BGType,
                BGAmount: item.BGAmount,
                Currency: item.Currency,
                ExpiryDate: formatDate(item.ExpiryDate),
                ClaimDate: formatDate(item.ClaimDate),
                Status: item.Status
            })
                .join(" ")
                .toLowerCase()
                .includes(lowerSearch)
        );

        setFilteredData(filtered);
        setCurrentPage(1);
    }, [searchTerm, listData]);


    // 🔹 Pagination
    const handlePageChange = (page: number) => {
        if (page >= 1 && page <= totalPages) {
            setCurrentPage(page);
        }
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
                        <h2> Deviation Initiator </h2>
                    </div>
                </div>
            </div>
            <div className='col-md-12 px-2 d-flex justify-content-between align-items-center flex-wrap'>
                <div>
                    <input type="text" placeholder="Search..."
                        className="form-control" style={{ width: "250px" }}
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                    />
                </div>
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
                                    <th className="px-4 py-2">Supplier Name</th>
                                    <th className="px-4 py-2">Batch / Lot No </th>
                                    <th className="px-4 py-2">Plant Name</th>
                                    <th className="px-4 py-2">Requester Name</th>
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
                                        <td className="px-4 py-2">{item.SupplierName}</td>
                                        <td className="px-4 py-2">{item.BatchNo}</td>
                                        <td className="px-4 py-2">{item.PlantName}</td>
                                        <td className="px-4 py-2">{item.RequestorName}</td>
                                        <td className="px-4 py-2">{item.DeviationType}</td>
                                        <td className="px-4 py-2">{item.Status}</td>
                                        <td className="px-4 py-2">
                                            <>
                                                <Link to={`/DeviationApproverForm/${item.Id}`}>
                                                    <img src={Edit} width={15} />
                                                </Link>
                                                &nbsp;&nbsp;
                                                <Link to={`/DeviationViewForm/${item.Id}`}>
                                                    <img src={View} width={15} />
                                                </Link>
                                            </>
                                        </td>

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

export default COODashboard;
