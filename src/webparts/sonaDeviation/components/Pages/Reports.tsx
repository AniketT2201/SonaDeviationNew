import * as React from "react";
import { useState, useEffect } from "react";
import { useHistory, Link } from "react-router-dom";
import { ISonaDeviationProps } from "../ISonaDeviationProps";
import "./Css/InitiatorDashboard.scss"
import SPCRUDOPS from "../../service/BAL/spcrud";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import {
    Chart as ChartJS,
    ArcElement,
    Tooltip,
    Legend,
} from "chart.js";

import { Doughnut } from "react-chartjs-2";
import ChartDataLabels from "chartjs-plugin-datalabels";

// ADD THIS PLUGIN ABOVE COMPONENT OR BEFORE RETURN
const centerTextPlugin = {
    id: "centerText",
    beforeDraw: (chart: any) => {
        const { width, height, ctx } = chart;
        ctx.restore();
        const dataset = chart.data.datasets[0];
        const total = dataset.data.reduce(
            (sum: number, item: any) =>
                sum + Number(item.value || 0),
            0
        );

        ctx.font = "bold 28px sans-serif";
        ctx.textBaseline = "middle";
        ctx.fillStyle = "#222";

        const text = total.toString();
        const textX =
            Math.round(
                (width - ctx.measureText(text).width) / 2
            );
        const textY = height / 2;

        ctx.fillText(text, textX, textY);
        ctx.save();
    },
};

ChartJS.register(
    ArcElement,
    Tooltip,
    Legend,
    ChartDataLabels,
    centerTextPlugin
);

export const Reports: React.FC<ISonaDeviationProps> = (props: ISonaDeviationProps) => {
    const history = useHistory();
    const [listData, setListData] = useState<any[]>([]);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [loading, setLoading] = useState(true);
    const [scope, setScope] = useState<string>("All");
    const [selectedDeviationType, setSelectedDeviationType] = useState<string>("");
    const [selectedStatus, setSelectedStatus] = useState<string>("");
    const [filters, setFilters] = useState({
        PartNo: "",
        DeviationNo: "",
        SupplierName: "",
        BatchNo: "",
        PlantName: "",
        RequestorName: "",
        DeviationType: "",
        Status: ""
    });

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

    React.useEffect(() => {
        sp.setup({
            spfxContext: props.currentSPContext
        });
    }, []);

    //Latest list data
    const GetListData = async () => {
        setLoading(true);
        const spCrudOps = await SPCRUDOPS();
        // 🔹 Get logged-in user
        const loggedInUser = await sp.web.currentUser();
        const parentItems = await spCrudOps.getData(
            "DeviationDetails",
            "*,ID,PartNo,DeviationNo,SupplierName,BatchNo,PlantName,ProductionQTY,RequestorName,Author/ID,Author/Title",
            "Author",
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

    const scopedData = React.useMemo(() => {
        if (scope === "My") {
            return listData.filter((item: any) => {
                return (
                    item.Author?.Id === props.id ||
                    item.RequestorName === props.userDisplayName
                );
            });
        }
        return listData;
    }, [scope, listData]);

    useEffect(() => {
        GetListData();
    }, []);

    // PARENT PIE -> Deviation Type
    const deviationTypeCounts: any = {};
    scopedData.forEach((item: any) => {
        const type = item.DeviationType || "Unknown";
        if (!deviationTypeCounts[type]) {
            deviationTypeCounts[type] = 0;
        }
        deviationTypeCounts[type]++;
    });
    const deviationLabels = Object.keys(deviationTypeCounts);
    const deviationValues: number[] = Object.values(deviationTypeCounts).map((v: any) => Number(v));
    const totalDeviation = deviationValues.reduce(
        (a: any, b: any) => Number(a) + Number(b),
        0
    );
//     const colors = [
//     "#0F766E", // rich teal
//     "#7C3AED", // premium violet
//     "#B45309", // elegant amber
//     "#BE123C", // sophisticated rose
//     "#374151", // slate gray
//     "#0EA5E9", // modern sky blue
//     "#15803D", // professional green
//     "#9333EA", // luxury purple
//     "#C2410C", // burnt orange
//     "#2A9D8F", // muted emerald
//     "#E9C46A", // luxury sand
//     "#F4A261", // warm peach
//     "#E76F51", // terracotta
//     "#5E548E", // royal muted purple
//     "#6B705C", // olive slate
//     "#B5838D", // dusty rose
//     "#264653", // charcoal teal
// ];
const colors = [
    "#355070", // deep steel blue
    "#6D597A", // muted royal purple
    "#B56576", // dusty rose
    "#E56B6F", // soft coral
    "#EAAC8B", // warm peach
    "#7F8C8D", // elegant gray
];
    const deviationChartData = {
        datasets: [

            // INNER RING
            {
                label: "Deviation Type",
                data: deviationLabels.map((label, index) => ({
                    label,
                    value: deviationValues[index],
                })),
                parsing: {
                    key: "value",
                },
                backgroundColor: colors,
                hoverOffset: 18,
                borderWidth: 2,
                borderColor: "#fff",
                cutout: "35%",
            },
        ],
    };
    const deviationOptions: any = {
        responsive: true,
        plugins: {
            legend: {
                display: false,
            },
            datalabels: {
                display: false,
            },
            tooltip: {
                callbacks: {
                    label: function (context: any) {
                        const dataset = context.dataset.label;
                        const label = context.raw.label;
                        const value = Number(context.raw.value);
                        const total = context.dataset.data.reduce(
                            (a: number, b: any) => a + Number(b.value),
                            0
                        );
                        const percentage = (
                            (value / total) * 100
                        ).toFixed(1);
                        return `${dataset} - ${label}: ${value} (${percentage}%)`;
                    },
                },
            },
        },
        onClick: (event: any, elements: any) => {
            if (elements.length > 0) {
                const index = elements[0].index;
                const clickedLabel = deviationLabels[index];
                setSelectedDeviationType(clickedLabel);
            }
        },
    };

    // CHILD PIE -> Deviation Type
    // =========================
    // CHILD CHART 1 -> PLANTS
    // =========================
    const deviationChildData = scopedData.filter(
        (item: any) => item.DeviationType === selectedDeviationType
    );
    const plantCounts: any = {};

    deviationChildData.forEach((item: any) => {
        const plant = item.PlantName || "Unknown";
        if (!plantCounts[plant]) {
            plantCounts[plant] = 0;
        }
        plantCounts[plant]++;
    });

    const plantLabels = Object.keys(plantCounts);
    const plantValues: number[] = Object.values(plantCounts).map((v: any) => Number(v));

    const plantChartData = {
        datasets: [
            {
                label: "Plants",
                data: plantLabels.map((label, index) => ({
                    label,
                    value: plantValues[index],
                })),
                parsing: {
                    key: "value",
                },
                backgroundColor: colors,
                hoverOffset: 18,
                borderWidth: 2,
                borderColor: "#fff",
                cutout: "35%",
            },
        ],
    };

    const plantOptions: any = {
        responsive: true,
        onClick: () => { },
        plugins: {
            legend: {
                display: false,
            },
            datalabels: {
                display: false,
            },
            tooltip: {
                callbacks: {
                    label: function (context: any) {
                        const label = context.raw.label;
                        const value = Number(context.raw.value);

                        return `${label}: ${value}`;
                    },
                },
            },
        },
    };



    // PARENT PIE -> Status Type 
    const statusCounts: any = {
        "Engineering Pending": 0,
        "Quality Pending": 0,
        "COO Pending": 0,
        "Plant Head Pending": 0,
        "Approved": 0,
        "Rejected": 0,
    };

    scopedData.forEach((item: any) => {
        // FINAL STATUS CHECK
        if (item.Status === "Approved") {
            statusCounts["Approved"]++;
            return;
        }
        if (item.Status === "Rejected") {
            statusCounts["Rejected"]++;
            return;
        }
        // WORKFLOW HISTORY CHECK
        if (item.WorkflowHistory) {
            try {
                const history = JSON.parse(item.WorkflowHistory);
                if (history.length > 0) {
                    const lastEntry = history[history.length - 1];
                    const currentStatus = lastEntry.CurrentStatus || "";

                    // ENGINEERING
                    if (currentStatus.includes("Engineering")) {
                        statusCounts["Engineering Pending"]++;
                    }
                    // QUALITY
                    else if (currentStatus.includes("Quality")) {
                        statusCounts["Quality Pending"]++;
                    }
                    // COO
                    else if (currentStatus.includes("COO")) {
                        statusCounts["COO Pending"]++;
                    }
                    // PLANT HEAD
                    else if (currentStatus.includes("Plant Head")) {
                        statusCounts["Plant Head Pending"]++;
                    }
                }
            } catch (err) {
                console.log(err);
            }
        }
    });
    const statusLabels = Object.keys(statusCounts);
    const statusValues: number[] = Object.values(statusCounts).map((v: any) => Number(v));
    const totalStatus = statusValues.reduce(
        (a: any, b: any) => Number(a) + Number(b),
        0
    );
    const statusChartData = {
        datasets: [
            // INNER RING
            {
                label: "Status",
                data: statusLabels.map((label, index) => ({
                    label,
                    value: statusValues[index],
                })),
                parsing: {
                    key: "value",
                },
                backgroundColor: colors,
                hoverOffset: 18,
                borderWidth: 2,
                borderColor: "#fff",
                cutout: "35%",
            },
        ],
    };
    const statusOptions: any = {
        responsive: true,
        plugins: {
            legend: {
                display: false,
            },
            datalabels: {
                display: false,
            },
            tooltip: {
                callbacks: {
                    label: function (context: any) {
                        const dataset = context.dataset.label;
                        const label = context.raw.label;
                        const value = Number(context.raw.value);
                        const total = context.dataset.data.reduce(
                            (a: number, b: any) => a + Number(b.value),
                            0
                        );
                        const percentage = (
                            (value / total) * 100
                        ).toFixed(1);
                        return `${dataset} - ${label}: ${value} (${percentage}%)`;
                    },
                },
            },
        },
        onClick: (event: any, elements: any) => {
            if (elements.length > 0) {
                const index = elements[0].index;
                const clickedLabel = statusLabels[index];
                setSelectedStatus(clickedLabel);
            }
        },
    };

    // CHILD PIE -> Status Type
    // =========================
    // CHILD CHART 2 -> ACTIONS
    // =========================
    const statusChildData = scopedData.filter((item: any) => item.Status === selectedStatus);
    const actionCounts: any = {};

    statusChildData.forEach((item: any) => {
        if (item.WorkflowHistory) {
            try {
                const history = JSON.parse(item.WorkflowHistory);
                history.forEach((h: any) => {
                    const action = h.ActionTaken || "Unknown";
                    if (!actionCounts[action]) {
                        actionCounts[action] = 0;
                    }
                    actionCounts[action]++;
                });
            } catch (err) {
                console.log(err);
            }
        }
    });

    const actionLabels = Object.keys(actionCounts);
    const actionValues: number[] = Object.values(actionCounts).map((v: any) => Number(v));

    const actionChartData = {
        datasets: [
            {
                label: "Actions",
                data: actionLabels.map((label, index) => ({
                    label,
                    value: actionValues[index],
                })),
                parsing: {
                    key: "value",
                },
                backgroundColor: colors,
                hoverOffset: 18,
                borderWidth: 2,
                borderColor: "#fff",
                cutout: "35%",
            },
        ],
    };

    const actionOptions: any = {
        responsive: true,
        plugins: {
            legend: {
                display: false,
            },
            datalabels: {
                display: false,
            },
            tooltip: {
                callbacks: {
                    label: function (context: any) {
                        const label = context.raw.label;
                        const value = Number(context.raw.value);

                        return `${label}: ${value}`;
                    },
                },
            },
        },
    };

    return (
        <div className="min-h-screen bg-gray-100">
            {/* Header */}
            <div className="p-4">
                <div
                    style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        marginBottom: "20px",
                    }}
                >
                    <h2
                        style={{
                            margin: 0,
                            fontSize: "22px",
                            fontWeight: 600,
                        }}
                    >
                        Deviation Reports
                    </h2>
                    <div>
                        <label
                            style={{
                                marginRight: "10px",
                                fontWeight: 500,
                            }}
                        >
                            Scope:
                        </label>
                        <select
                            value={scope}
                            onChange={(e) => setScope(e.target.value)}
                            style={{
                                padding: "8px 14px",
                                borderRadius: "8px",
                                border: "1px solid #ccc",
                                outline: "none",
                                cursor: "pointer",
                            }}
                        >
                            <option value="All">All Reports</option>
                            <option value="My">My Reports</option>
                        </select>
                    </div>
                </div>
                <div
                    style={{
                        display: "flex",
                        flexDirection: "column",
                        gap: "25px",
                    }}
                >
                    {/* ================= ROW 1 ================= */}
                    <div
                        style={{
                            display: "flex",
                            gap: "20px",
                            justifyContent: "center",
                            alignItems: "stretch",
                            flexWrap: "wrap",
                        }}
                    >
                        {/* PARENT CHART */}
                        <div
                            style={{
                                width: "420px",
                                background: "#fff",
                                borderRadius: "16px",
                                padding: "20px",
                                boxShadow: "0 4px 14px rgba(0,0,0,0.08)",
                            }}
                        >
                            <h3
                                style={{
                                    textAlign: "center",
                                    marginBottom: "15px",
                                    fontSize: "18px",
                                    fontWeight: 600,
                                }}
                            >
                                Deviation Type Report
                            </h3>
                            <div style={{ height: "300px" }}>
                                <Doughnut
                                    data={deviationChartData}
                                    options={{
                                        ...deviationOptions,
                                        maintainAspectRatio: false,
                                    }}
                                />
                            </div>
                            <div
                                style={{
                                    marginTop: "20px",
                                    display: "flex",
                                    flexWrap: "wrap",
                                    gap: "12px",
                                    justifyContent: "center",
                                }}
                            >
                                {deviationLabels.map((label, index) => (
                                    <div
                                        key={index}
                                        style={{
                                            display: "flex",
                                            alignItems: "center",
                                            gap: "6px",
                                            fontSize: "13px",
                                            cursor: "pointer",
                                        }}
                                        onClick={() => setSelectedDeviationType(label)}
                                    >
                                        <div
                                            style={{
                                                width: "14px",
                                                height: "14px",
                                                borderRadius: "50%",
                                                background: colors[index % colors.length],
                                            }}
                                        />
                                        <span
                                            style={{
                                                fontWeight:
                                                    selectedDeviationType === label
                                                        ? 700
                                                        : 400,
                                            }}
                                        >
                                            {label}
                                        </span>
                                    </div>
                                ))}
                            </div>
                        </div>
                        {/* CHILD CHART */}
                        <div
                            style={{
                                width: "420px",
                                background: "#fff",
                                borderRadius: "16px",
                                padding: "20px",
                                boxShadow: "0 4px 14px rgba(0,0,0,0.08)",
                            }}
                        >
                            <h3
                                style={{
                                    textAlign: "center",
                                    marginBottom: "15px",
                                    fontSize: "18px",
                                    fontWeight: 600,
                                }}
                            >
                                Plant Distribution ({selectedDeviationType})
                            </h3>

                            <div style={{ height: "300px" }}>
                                <Doughnut
                                    data={plantChartData}
                                    options={{
                                        ...plantOptions,
                                        maintainAspectRatio: false,
                                    }}
                                />
                            </div>
                            <div
                                style={{
                                    marginTop: "20px",
                                    display: "flex",
                                    flexWrap: "wrap",
                                    gap: "12px",
                                    justifyContent: "center",
                                }}
                            >
                                {plantLabels.map((label, index) => (
                                    <div
                                        key={index}
                                        style={{
                                            display: "flex",
                                            alignItems: "center",
                                            gap: "6px",
                                            fontSize: "13px",
                                        }}
                                    >
                                        <div
                                            style={{
                                                width: "14px",
                                                height: "14px",
                                                borderRadius: "50%",
                                                background: colors[index % colors.length],
                                            }}
                                        />
                                        <span>{label}</span>
                                    </div>
                                ))}
                            </div>
                        </div>
                    </div>
                    {/* ================= ROW 2 ================= */}
                    <div
                        style={{
                            display: "flex",
                            gap: "20px",
                            justifyContent: "center",
                            alignItems: "stretch",
                            flexWrap: "wrap",
                        }}
                    >
                        {/* PARENT STATUS */}
                        <div
                            style={{
                                width: "420px",
                                background: "#fff",
                                borderRadius: "16px",
                                padding: "20px",
                                boxShadow: "0 4px 14px rgba(0,0,0,0.08)",
                            }}
                        >
                            <h3
                                style={{
                                    textAlign: "center",
                                    marginBottom: "15px",
                                    fontSize: "18px",
                                    fontWeight: 600,
                                }}
                            >
                                Status Report
                            </h3>
                            <div style={{ height: "300px" }}>
                                <Doughnut
                                    data={statusChartData}
                                    options={{
                                        ...statusOptions,
                                        maintainAspectRatio: false,
                                    }}
                                />
                            </div>
                            <div
                                style={{
                                    marginTop: "20px",
                                    display: "flex",
                                    flexWrap: "wrap",
                                    gap: "12px",
                                    justifyContent: "center",
                                }}
                            >
                                {statusLabels.map((label, index) => (
                                    <div
                                        key={index}
                                        style={{
                                            display: "flex",
                                            alignItems: "center",
                                            gap: "6px",
                                            fontSize: "13px",
                                            cursor: "pointer",
                                        }}
                                        onClick={() => setSelectedStatus(label)}
                                    >
                                        <div
                                            style={{
                                                width: "14px",
                                                height: "14px",
                                                borderRadius: "50%",
                                                background: colors[index % colors.length],
                                            }}
                                        />
                                        <span
                                            style={{
                                                fontWeight:
                                                    selectedStatus === label
                                                        ? 700
                                                        : 400,
                                            }}
                                        >
                                            {label}
                                        </span>
                                    </div>
                                ))}
                            </div>
                        </div>
                        {/* CHILD ACTION CHART */}
                        {/* <div
                            style={{
                                width: "420px",
                                background: "#fff",
                                borderRadius: "16px",
                                padding: "20px",
                                boxShadow: "0 4px 14px rgba(0,0,0,0.08)",
                            }}
                        >
                            <h3
                                style={{
                                    textAlign: "center",
                                    marginBottom: "15px",
                                    fontSize: "18px",
                                    fontWeight: 600,
                                }}
                            >
                                Action Taken ({selectedStatus})
                            </h3>
                            <div style={{ height: "300px" }}>
                                <Doughnut
                                    data={actionChartData}
                                    options={{
                                        ...actionOptions,
                                        maintainAspectRatio: false,
                                    }}
                                />
                            </div>
                            <div
                                style={{
                                    marginTop: "20px",
                                    display: "flex",
                                    flexWrap: "wrap",
                                    gap: "12px",
                                    justifyContent: "center",
                                }}
                            >
                                {actionLabels.map((label, index) => (
                                    <div
                                        key={index}
                                        style={{
                                            display: "flex",
                                            alignItems: "center",
                                            gap: "6px",
                                            fontSize: "13px",
                                        }}
                                    >
                                        <div
                                            style={{
                                                width: "14px",
                                                height: "14px",
                                                borderRadius: "50%",
                                                background: colors[index % colors.length],
                                            }}
                                        />
                                        <span>{label}</span>
                                    </div>
                                ))}
                            </div>
                        </div> */}
                    </div>
                </div>
            </div>
        </div>
    );
};

export default Reports;
