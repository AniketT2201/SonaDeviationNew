

import * as React from 'react';
import { useState, useEffect } from 'react';
import type { ISonaDeviationProps } from '../ISonaDeviationProps';
import { TextField, DefaultButton, Dropdown, IDropdownOption } from '@fluentui/react';
import { Link, useHistory, useLocation } from "react-router-dom";
import logo from "../../assets/sona-comstarlogo.png";
import "./Css/NewRequest.scss";

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import { useParams } from "react-router-dom";
import { set } from '@microsoft/sp-lodash-subset/lib/index';

import html2canvas from "html2canvas";
import { jsPDF } from "jspdf"

interface IRouteParams {
    id: string;
}

export const DeviationViewForm = (props: ISonaDeviationProps) => {
    {
        const { id } = useParams<IRouteParams>();
        const history = useHistory();
        const location = useLocation<any>();

        const [ApprovalMatrixdata, setApprovalMatrix] = React.useState<any[]>([]);

        const [WorkflowHistorydata, setWorkflowHistory] = React.useState<any[]>([]);

        const [Stage, setStage] = React.useState(0);
        const [isSubmitting, setIsSubmitting] = useState(false);

        const [PartNo, setPartNo] = React.useState("");
        const [DeviationNo, setDeviationNo] = React.useState("");
        const [ConcernedDepartment, setConcernedDepartment] = React.useState("");
        const [BatchNo, setBatchNo] = React.useState("");
        const [PlantName, setPlantName] = React.useState("");
        const [ProductionQTY, setProductionQTY] = React.useState("");
        const [RequestorName, setRequestorName] = React.useState("");
        const [ProductionDate, setProductionDate] = React.useState("");
        const [DeviationType, setDeviationType] = React.useState("");


        const [Remark, setRemark] = React.useState("");
        const [Functional, setFunctional] = React.useState("");
        const [Fitment, setFitment] = React.useState("");
        const [SpecialNote, setSpecialNote] = React.useState("");
        const [DeviationRecommended, setDeviationRecommended] = React.useState("");
        const [Name, setName] = React.useState("");
        const [QualityProductionDate, setQualityProductionDate] = React.useState("");
        const [CDate, setDate] = React.useState("");

        const [EngineeringRemark, setEngineeringRemark] = React.useState("");
        const [EngineeringFunctional, setEngineeringFunctional] = React.useState("");
        const [EngineeringFitment, setEngineeringFitment] = React.useState("");
        const [EngineeringSpecialNote, setEngineeringSpecialNote] = React.useState("");
        const [EngineeringDeviationRecommended, setEngineeringDeviationRecommended] = useState<number | undefined>(undefined);
        const [EngineeringName, setEngineeringName] = React.useState("");
        const [EngineeringProductionDate, setEngineeringProductionDate] = React.useState("");
        const [EngineeringDate, setEngineeringDate] = React.useState("");

        const [COORemark, setCOORemark] = React.useState("");

        const [PlantHeadRemark, setPlantHeadRemark] = React.useState("");
        const [PlantHeadDeviationAccepted, setPlantHeadDeviationAccepted] = useState<number | undefined>(undefined);
        const [PlantHeadProductionDate, setPlantHeadProductionDate] = React.useState("");
        const [PlantHeadDate, setPlantHeadDate] = React.useState("");
        const [PlantHeadEffectiveness, setPlantHeadEffectiveness] = React.useState("");


        const [Details, setDetails] = React.useState<any[]>([]);

        const [QualityApprover, setQualityApprover] = React.useState(false);
        const [EngineeringApprover, setEngineeringApprover] = React.useState(false);
        const [PlantHeadApprover, setPlantHeadApprover] = React.useState(false);
        const [COOApprover, setCOOApprover] = React.useState(false);

        const [Status, setStatus] = React.useState("");
        const [SpecialApproval, setSpecialApproval] = React.useState("");

        const [DeviationAttachments, setDeviationAttachments] = React.useState<any[]>([]);

        const [WorkflowJSX, setWorkflowJSX] = React.useState(null);

        const [Comments, setComments] = React.useState("");
        const [selectedRiskAssessment, setSelectedRiskAssessment] = React.useState("");
        const [RiskAssessmentDescription, setRiskAssessmentDescription] = React.useState("");

        const pdfRef = React.useRef<HTMLDivElement>(null);

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
            if (ApprovalMatrixdata && ApprovalMatrixdata.length > 0 && WorkflowHistorydata) {
                displayWorkflow();
            }
        }, [ApprovalMatrixdata, WorkflowHistorydata, Stage])

        const displayWorkflow = () => {
            const wf: JSX.Element[] = [];

            // const _wf = approverJson.filter((item) => item.required === true);
            let isActive;
            let notActive = false;

            // Get latest workflow history entry
            const lastHistory = WorkflowHistorydata && WorkflowHistorydata.length > 0
                ? WorkflowHistorydata[WorkflowHistorydata.length - 1]
                : null;

            // Check if rejected
            const isRejected =
                lastHistory &&
                lastHistory.CurrentStatus === "Rejected";

            // Find rejected role
            let rejectedRole = "";

            if (isRejected && lastHistory.ActionTaken) {
                rejectedRole = lastHistory.ActionTaken.split(" ")[0];
            }
            // Find rejected role index
            const rejectedIndex = ApprovalMatrixdata
                .filter(m => m.required === true)
                .findIndex(m => m.Role === rejectedRole);
            ApprovalMatrixdata.filter(m => m.required === true).forEach((m, i) => {
                // =========================
                // REJECTED LOGIC
                // =========================
                if (isRejected) {
                    if (m.Role === rejectedRole) {
                        // Rejected approver → RED
                        isActive = "rejected";
                    }
                    else if (i < rejectedIndex) {
                        // Completed before rejection → GREEN
                        isActive = "beforeactiveApprover";
                    }
                    else {
                        // Pending after rejection → YELLOW
                        isActive = "overrideStage";
                    }
                }
                // =========================
                // NORMAL APPROVAL FLOW
                // =========================
                else {
                    if (notActive === false && Stage !== 99) {
                        if (Stage === i) {
                            isActive = "activeApprover";
                            notActive = true;
                        }
                        else {
                            isActive = "beforeactiveApprover";
                        }
                    }
                    else {
                        isActive = "overrideStage";
                    }
                }
                // wf.push(
                //     <ul className="main-menu">
                //         <li className={`${m.Role} ${isActive}`.trim()}>
                //             {m.User}
                //         </li>
                //     </ul>
                // );
                wf.push(
                    <ul className="main-menu" key={i}>
                        <li className={`${m.Role} ${isActive}`.trim()}>
                            {m.User}
                        </li>
                    </ul>
                );
            });

            setWorkflowJSX(wf);
        };

        React.useEffect(() => {
            sp.setup({
                spfxContext: props.currentSPContext
            });

            if (id) {
                loadItem(Number(id));
            }

        }, [id]);

        const filterdata = async (item, roletype) => {
            //if (item.NextApprover?.ID !== props.id) return false;

            let matrix: any[] = [];

            try {
                matrix = JSON.parse(item.ApprovalMatrix || "[]");
            } catch (e) {
                console.error("Invalid ApprovalMatrix JSON", item);
                return false;
            }

            const engineeringStep = matrix.find(
                (m: any) => m.Role === roletype && m.required === true
            );

            return engineeringStep?.Stage === item.Stage;
        }

        const loadItem = async (itemId: number) => {
            try {
                const item = await sp.web.lists
                    .getByTitle("DeviationDetails").items.getById(itemId)
                    .select("ID", "PartNo", "DeviationNo", "SupplierName", "BatchNo",
                        "PlantName", "ProductionQTY", "RequestorName", "ProductionDate",
                        "Details", "ApprovalMatrix", "Stage", "WorkflowHistory",
                        "Remark", "Functional", "Fitment", "SpecialNote", "DeviationRecommended", "QualityName", "QualityProductionDate", "QualityDate",
                        "EngineeringRemark", "EngineeringFunctional", "EngineeringFitment", "EngineeringSpecialNote", "EngineeringDeviationRecommended",
                        "EngineeringName", "EngineeringProductionDate", "EngineeringDate", "DeviationType", "Status", "SpecialApproval",
                        "PlantHeadRemark", "PlantHeadDeviationAccepted", "PlantHeadProductionDate", "COORemark",
                        "PlantHeadDate", "PlantHeadEffectiveness", "AttachmentFiles", "Comments", "RiskAssessment", "RiskAssessmentDescription"
                    ).expand("AttachmentFiles")();
                setPartNo(item.PartNo || "");
                setDeviationNo(item.DeviationNo || "");
                setConcernedDepartment(item.SupplierName || "");
                setBatchNo(item.BatchNo || "");
                setPlantName(item.PlantName || "");
                setProductionQTY(item.ProductionQTY || "");
                setRequestorName(item.RequestorName || "");
                setProductionDate(item.ProductionDate || "");
                setDeviationType(item.DeviationType || "");
                setStage(item.Stage || "");
                setDeviationAttachments(item.AttachmentFiles || []);
                setStatus(item.Status || "");

                setRemark(item.Remark || "");
                setFunctional(item.Functional || "");
                setFitment(item.Fitment || "");
                setSpecialNote(item.SpecialNote || "");
                setDeviationRecommended(item.DeviationRecommended || "");
                setName(item.QualityName || "");
                setQualityProductionDate(item.QualityProductionDate || "");
                setDate(item.QualityDate || "");

                setCOORemark(item.COORemark || "");
                setSpecialApproval(item.SpecialApproval || "");

                setEngineeringRemark(item.EngineeringRemark || "");
                setEngineeringFunctional(item.EngineeringFunctional || "");
                setEngineeringFitment(item.EngineeringFitment || "");
                setEngineeringSpecialNote(item.EngineeringSpecialNote || "");
                setEngineeringDeviationRecommended(item.EngineeringDeviationRecommended || "");
                setEngineeringName(item.EngineeringName || "");
                setEngineeringProductionDate(item.EngineeringProductionDate || "");
                setEngineeringDate(item.EngineeringDate || "");

                setPlantHeadRemark(item.PlantHeadRemark || "");
                setPlantHeadDeviationAccepted(item.PlantHeadDeviationAccepted || "");
                setPlantHeadProductionDate(item.PlantHeadProductionDate || "");
                setPlantHeadDate(item.PlantHeadDate || "");
                setPlantHeadEffectiveness(item.PlantHeadEffectiveness || "");
                setComments(item.Comments || "");
                setSelectedRiskAssessment(item.RiskAssessment || "");
                setRiskAssessmentDescription(item.RiskAssessmentDescription || "");


                let parsedDetails = [];
                if (item.Details) {
                    try {
                        parsedDetails = JSON.parse(item.Details);
                    } catch (parseError) {
                        console.error("Error parsing ApprovalMatrix JSON:", parseError);
                    }
                }
                setDetails(parsedDetails);

                let parsedApprovalMatrix = [];
                if (item.ApprovalMatrix) {
                    try {
                        parsedApprovalMatrix = JSON.parse(item.ApprovalMatrix);
                    } catch (parseError) {
                        console.error("Error parsing ApprovalMatrix JSON:", parseError);
                    }
                }

                setApprovalMatrix(parsedApprovalMatrix);


                let parsedWorkflowHistory = [];
                if (item.WorkflowHistory) {
                    try {
                        parsedWorkflowHistory = JSON.parse(item.WorkflowHistory);
                    } catch (parseError) {
                        console.error("Error parsing WorkflowHistory JSON:", parseError);
                    }
                }

                setWorkflowHistory(parsedWorkflowHistory);

                let Roledata;
                let RoletypeArr = ["Quality", "Engineering", "PlantHead", "COO"];
                for (var i = 0; i < RoletypeArr.length; i++) {
                    let stage = await filterdata(item, RoletypeArr[i]);
                    if (item.Status === "Approved") {
                        setEngineeringApprover(true)
                        setCOOApprover(true)
                        setQualityApprover(true)
                        setPlantHeadApprover(true)
                    } else if (stage === true) {

                        // if (RoletypeArr[i] === "Engineering") { setEngineeringApprover(true) };
                        // if (RoletypeArr[i] === "COO") { setCOOApprover(true); };
                        // if (RoletypeArr[i] === "Quality") { setQualityApprover(true); };
                        // if (RoletypeArr[i] === "PlantHead") { setPlantHeadApprover(true); };

                        if (RoletypeArr[i] === "Engineering") { setEngineeringApprover(true) }
                        if (RoletypeArr[i] === "COO") { setCOOApprover(true) }
                        if (RoletypeArr[i] === "Quality") { setQualityApprover(true) }
                        if (RoletypeArr[i] === "PlantHead") { setPlantHeadApprover(true) }
                    }

                }
            } catch (error) {
                console.error("Error loading BG item:", error);
            }
        };

        const getFullUrl = (relativeUrl: string) => {
            const baseUrl = props.currentSPContext.pageContext.web.absoluteUrl.split('/sites')[0];
            return baseUrl + relativeUrl;
        };

        const showOthersColumn = Details.some(r => r.others && r.others.trim() !== "");

        const isEngineeringDashboard = location.state?.fromEngineering || false;


        const generatePDF = async () => {

            if (!pdfRef.current) return;

            const element = pdfRef.current;

            // 🔥 FIX: Force stable width for deployment rendering
            const originalWidth = element.style.width;
            element.style.width = "1200px";

            const canvas = await html2canvas(element, {
                scale: 3,
                useCORS: true,
                allowTaint: true,
                backgroundColor: "#ffffff",
                logging: false,
                scrollX: 0,
                scrollY: -window.scrollY,

                onclone: (clonedDoc) => {

                    // =====================================================
                    // Heading SVG Conversion
                    // =====================================================

                    const headings =
                        clonedDoc.querySelectorAll(".heading1");

                    headings.forEach((heading) => {

                        const headingText =
                            heading.querySelector("label")?.textContent ||
                            heading.textContent ||
                            "";

                        heading.innerHTML = `
                <svg xmlns="http://www.w3.org/2000/svg"
                    width="290" height="35" viewBox="0 0 290 35">
                    <defs>
                        <linearGradient id="headingGradient" x1="0%" y1="0%" x2="100%" y2="0%">
                            <stop offset="0%" stop-color="#000000" />
                            <stop offset="100%" stop-color="#bea1a1" />
                        </linearGradient>
                    </defs>

                    <polygon points="0,0 174,0 290,35 0,35"
                        fill="url(#headingGradient)" />

                    <text x="15" y="22"
                        fill="white"
                        font-size="13"
                        font-family="Arial"
                        font-weight="bold">
                        ${headingText}
                    </text>
                </svg>
                `;

                        const h = heading as HTMLElement;

                        h.style.width = "290px";
                        h.style.height = "35px";
                        h.style.padding = "0";
                        h.style.background = "transparent";
                        h.style.clipPath = "none";
                    });


                    // =====================================================
                    // WORKFLOW FIX (IMPORTANT - DEPLOYMENT ISSUE FIX)
                    // =====================================================

                    const workflowContainers =
                        clonedDoc.querySelectorAll(".displayWF");

                    workflowContainers.forEach((container) => {

                        const displayWF = container as HTMLElement;

                        // 🔥 FORCE SINGLE LINE LAYOUT
                        displayWF.style.display = "flex";
                        displayWF.style.flexDirection = "row";
                        displayWF.style.flexWrap = "nowrap";
                        displayWF.style.alignItems = "center";
                        displayWF.style.justifyContent = "space-between";
                        displayWF.style.gap = "6px";
                        displayWF.style.padding = "10px";
                        displayWF.style.backgroundColor = "#333";
                        displayWF.style.width = "100%";
                        displayWF.style.boxSizing = "border-box";
                        displayWF.style.overflow = "hidden";

                        const menus =
                            displayWF.querySelectorAll(".main-menu");

                        const totalMenus = menus.length;
                        const containerWidth = 1200; // FIXED WIDTH FOR PDF
                        const gap = 6;

                        const menuWidth =
                            (containerWidth -
                                ((totalMenus - 1) * gap) -
                                20) / totalMenus;

                        menus.forEach((menu) => {

                            const htmlMenu = menu as HTMLElement;

                            htmlMenu.style.display = "flex";
                            htmlMenu.style.flex = "1 1 0";
                            htmlMenu.style.minWidth = "0";
                            htmlMenu.style.margin = "0";
                            htmlMenu.style.padding = "0";
                            htmlMenu.style.listStyle = "none";
                            htmlMenu.style.flexWrap = "nowrap";
                            htmlMenu.style.justifyContent = "center";

                            const li =
                                htmlMenu.querySelector("li") as HTMLElement;

                            if (!li) return;

                            const text =
                                li.textContent?.trim() || "";

                            let bgColor = "#555";

                            if (li.classList.contains("beforeactiveApprover")) {
                                bgColor = "#0A9A00";
                            }
                            else if (li.classList.contains("overrideStage")) {
                                bgColor = "#FFB233";
                            }
                            else if (li.classList.contains("activeApprover")) {
                                bgColor = "#ED7D31";
                            }
                            else if (li.classList.contains("rejected")) {
                                bgColor = "#ff0000";
                            }

                            const width = Math.max(menuWidth, 180);
                            const height = 50;

                            li.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg"
                        width="${width}" height="${height}"
                        viewBox="0 0 ${width} ${height}">
                        
                        <polygon points="
                            0,0
                            ${width - 20},0
                            ${width},25
                            ${width - 20},50
                            0,50
                            20,25"
                            fill="${bgColor}" />

                        <text x="${width / 2}" y="25"
                            fill="white"
                            font-size="13"
                            font-family="Arial"
                            font-weight="bold"
                            text-anchor="middle"
                            dominant-baseline="middle">
                            ${text}
                        </text>
                    </svg>
                    `;

                            li.style.width = "100%";
                            li.style.height = "50px";
                            li.style.display = "flex";
                            li.style.alignItems = "center";
                            li.style.justifyContent = "center";
                            li.style.whiteSpace = "nowrap";
                            li.style.overflow = "hidden";
                            li.style.padding = "0";
                            li.style.margin = "0";
                            li.style.background = "transparent";
                            li.style.clipPath = "none";
                        });
                    });
                }
            });

            // restore width
            element.style.width = originalWidth;

            const imgData = canvas.toDataURL("image/png", 1.0);

            const pdf = new jsPDF("p", "mm", "a4");

            const pdfWidth = 210;
            const pdfHeight = 297;

            const imgWidth = pdfWidth;

            const imgHeight = (canvas.height * imgWidth) / canvas.width;

            let heightLeft = imgHeight;
            let position = 0;

            pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight, undefined, "FAST");

            heightLeft -= pdfHeight;

            while (heightLeft > 0) {

                position = heightLeft - imgHeight;

                pdf.addPage();

                pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight, undefined, "FAST");

                heightLeft -= pdfHeight;
            }

            const pdfBlobUrl = pdf.output("bloburl");
            window.open(pdfBlobUrl, "_blank");
        };

        return (
            <div ref={pdfRef} className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>
                            <div className="bordered">
                                <a><img src={logo} /></a>
                                <h1>Deviation Viewmore </h1>
                            </div>
                            <div className='displayWF'>{WorkflowJSX}</div>
                            <div className='borderedbox'>
                                <div className="heading1">
                                    <label>Requestor Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label htmlFor="Part No" className='font'>Part No </label> : &nbsp;&nbsp;
                                            <label className='fonttext'>{PartNo}</label>
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="Deviation No" className='font'>Deviation No </label> : &nbsp;&nbsp;
                                            <label className='fonttext'>{DeviationNo}</label>
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="Concerned Department" className='font'>Concerned Department </label> : &nbsp;&nbsp;
                                            <label className='fonttext'>{ConcernedDepartment}</label>
                                        </div>
                                    </div>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label htmlFor="Batch / Lot No" className='font'>Batch / Lot No </label> : &nbsp;&nbsp;
                                            <label className='fonttext'>{BatchNo}</label>
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="Plant Name" className='font'>Plant Name </label> : &nbsp;&nbsp;
                                            <label className='fonttext'>{PlantName}</label>
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="Production QTY" className='font'>Production QTY </label> : &nbsp;&nbsp;
                                            <label className='fonttext'>{ProductionQTY}</label>
                                        </div>
                                    </div>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label htmlFor="Requester Name" className='font'>Requester Name </label> : &nbsp;&nbsp;
                                            <label className='fonttext'>{RequestorName}</label>
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="ProductionDate" className='font'>Production Date </label> : &nbsp;&nbsp;
                                            <label className='fonttext'>{formatDate(ProductionDate)}</label>
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="Deviation Type" className='font'>Deviation Type </label> : &nbsp;&nbsp;
                                            <label className='fonttext'>{DeviationType}</label>
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Upload Documents</label>
                                </div>
                                <div className="main-formcontainer">
                                    <div className='row mb-20'>
                                        <div className='col-md-3'>
                                            <label className='font'>Attachments </label>
                                            <div>

                                                {DeviationAttachments.length > 0 ? (
                                                    <a href="#" 
                                                      onClick={(e) => {
                                                        e.preventDefault();
                                                        window.open(getFullUrl(DeviationAttachments[0].ServerRelativeUrl), "_blank", "noopener,noreferrer");
                                                      }}
                                                    >
                                                        {(DeviationAttachments[0].FileName)}
                                                    </a>
                                                ) : (
                                                    <span>No attachment found.</span>
                                                )}
                                            </div>
                                        </div>
                                        {isEngineeringDashboard && (
                                            <div className='col-md-3'>
                                                <label htmlFor="Comments" className='font'>Comments </label>  <br></br>
                                                <div><label className='fonttext'>{Comments}</label></div>
                                            </div>
                                        )}
                                        <div className='col-md-3'>
                                            <label htmlFor="RiskAssessment" className='font'>Risk Assessment</label> <br></br>
                                            <label className='fonttext'>{selectedRiskAssessment}</label>
                                        </div>
                                        <div className='col-md-3'>
                                            <label htmlFor="RiskAssessmentDescription" className='font'>Risk Assessment Description </label> <br></br>
                                            <label className='fonttext'>{RiskAssessmentDescription}</label>
                                        </div>
                                    </div>
                                </div>

                                {/* 🔹 Table */}
                                <div className='RequestTable'>
                                    <div style={{ overflowX: "auto", width: "100%" }}>
                                        <table>
                                            {/* <colgroup>
                                                <col style={{ width: "60px" }} />
                                                <col style={{ width: "220px" }} />
                                                <col style={{ width: "220px" }} />
                                                {showOthersColumn && <col style={{ width: "220px" }} />}
                                                <col style={{ width: "180px" }} />
                                                <col style={{ width: "180px" }} />
                                                <col style={{ width: "150px" }} />
                                                <col style={{ width: "100px" }} />
                                                <col style={{ width: "250px" }} />
                                                <col style={{ width: "250px" }} />
                                            </colgroup> */}
                                            <colgroup>
                                                <col style={{ width: "45px" }} />
                                                <col style={{ width: "130px" }} />
                                                <col style={{ width: "140px" }} />
                                                {showOthersColumn && <col style={{ width: "220px" }} />}
                                                <col style={{ width: "120px" }} />
                                                <col style={{ width: "120px" }} />
                                                <col style={{ width: "135px" }} />
                                                <col style={{ width: "80px" }} />
                                                <col style={{ width: "240px" }} />
                                                <col style={{ width: "140px" }} />
                                            </colgroup>

                                            <thead>
                                                <tr>
                                                    <th>S.N</th>
                                                    <th>Parameters</th>
                                                    <th>Sub Parameters</th>
                                                    {showOthersColumn && <th>Others</th>}
                                                    <th>Specification</th>
                                                    <th>Observation</th>
                                                    <th>Processing Date</th>
                                                    <th>Quantity</th>
                                                    <th>Root Cause of Defects Reported</th>
                                                    <th>Corrective Action</th>
                                                </tr>
                                            </thead>
                                            <tbody>
                                                {Details && Details.length > 0 ? (
                                                    Details.map((row: any, index: number) => (
                                                        <tr key={index}>
                                                            <td>{index + 1}</td>
                                                            <td>{row.Parameters || ""}</td>
                                                            <td>{row.SubParameters || ""}</td>
                                                            {showOthersColumn && <td>{row.others || ""}</td>}
                                                            <td>{row.Specification || ""}</td>
                                                            <td>{row.Observation || ""}</td>
                                                            <td style={{ whiteSpace: "nowrap" }}>{formatDate(row.Date)}</td>
                                                            <td>{row.Quantity || ""}</td>
                                                            <td>{row.RootCause || ""}</td>
                                                            <td>{row.CorrectiveAction || ""}</td>
                                                        </tr>
                                                    ))
                                                ) : (
                                                    <tr>
                                                        <td colSpan={9} style={{ textAlign: "center" }}>
                                                            No Data Available
                                                        </td>
                                                    </tr>
                                                )}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>

                                {((COOApprover || QualityApprover || PlantHeadApprover) || (Status === 'Rejected' || Status === 'Send Back' || Status === "Pending for Approval")) && (
                                    <>
                                        <div className='heading1'>
                                            <label>Recommendation Engineering</label>
                                        </div>

                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'> Special Approval</label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{SpecialApproval}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Deviation Recommended</label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{EngineeringDeviationRecommended}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Functional </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{EngineeringFunctional}</label>
                                                </div>
                                            </div>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Remarks </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{EngineeringRemark}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Fitment </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{EngineeringFitment}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Special Note </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{EngineeringSpecialNote}</label>
                                                </div>
                                            </div>
                                            {/* <div className='row mb-20'>          
                                                <div className='col-md-4'>
                                                    <label className='font'>Name </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{EngineeringName}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Production Date </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{formatDate(EngineeringProductionDate)}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Date </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{formatDate(EngineeringDate)}</label>
                                                </div>
                                            </div> */}
                                        </div>
                                    </>
                                )}

                                {(((QualityApprover || PlantHeadApprover) && SpecialApproval === "Yes") || ((Status === 'Rejected' || Status === 'Send Back' || Status === "Pending for Approval") && COORemark)) && (
                                    <div>
                                        <div className='heading1'>
                                            <label>Recommendation COO</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Remarks </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{COORemark}</label>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                )}

                                {((PlantHeadApprover && SpecialApproval === "No") || ((Status === 'Rejected' || Status === 'Send Back' || Status === "Pending for Approval") && DeviationRecommended && Functional && Remark && Fitment && SpecialNote)) && (
                                    <>
                                        <div className='heading1'>
                                            <label>Recommendation Quality</label>
                                        </div>

                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Deviation Recommended</label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{DeviationRecommended}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Functional </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{Functional}</label>
                                                </div>
                                                <div className='col-md-4 '>
                                                    <label className='font'>Remarks </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{Remark}</label>
                                                </div>
                                            </div>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Fitment </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{Fitment}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Special Note </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{SpecialNote}</label>
                                                </div>
                                            </div>
                                        </div>
                                    </>
                                )}

                                {PlantHeadApprover && SpecialApproval === "No" && (
                                    <>
                                        <div className='heading1'>
                                            <label>Recommendation PlantHead</label>
                                        </div>

                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Deviation Accepted </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{PlantHeadDeviationAccepted}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Remarks </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{PlantHeadRemark}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Effectiveness</label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{PlantHeadEffectiveness}</label>
                                                </div>
                                            </div>
                                            {/* <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Date </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{formatDate(PlantHeadDate)}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Production Date </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{formatDate(PlantHeadProductionDate)}</label>
                                                </div> 
                                            </div> */}
                                        </div>
                                    </>
                                )}


                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Work Flow History</label>
                                </div>
                                <div className="main-formcontainer">
                                    <div className='Workflowbox'>
                                        {WorkflowHistorydata && WorkflowHistorydata.length > 0 ? (
                                            <table className="workflow-table" style={{ width: '100%', borderCollapse: 'collapse' }}>
                                                <thead>
                                                    <tr>
                                                        <th style={{ padding: '8px', textAlign: 'left' }}>Action By</th>
                                                        <th style={{ padding: '8px', textAlign: 'left' }}>Action Taken</th>
                                                        <th style={{ padding: '8px', textAlign: 'left' }}>Current Status</th>
                                                        <th style={{ padding: '8px', textAlign: 'left' }}>Date</th>
                                                        <th style={{ padding: '8px', textAlign: 'left' }}>Comment</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {WorkflowHistorydata.map((h: any, idx: number) => (
                                                        <tr key={idx}>
                                                            <td style={{ padding: '8px' }}>{h.CurrentApprover || ''}</td>
                                                            <td style={{ padding: '8px' }}>{h.ActionTaken || ''}</td>
                                                            <td style={{ padding: '8px' }}>{h.CurrentStatus || ''}</td>
                                                            <td style={{ padding: '8px' }}>{h.Date || ''}</td>
                                                            <td style={{ padding: '8px' }}>{h.Comment || ''}</td>
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        ) : (
                                            <p>No workflow history</p>
                                        )}
                                    </div>
                                </div>

                                <div className='row my-3'>
                                    <div className='col-md-12'>
                                        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "10px" }}>
                                            <a className="print-btn" onClick={generatePDF}>
                                                Print
                                            </a>
                                            <Link to="#" onClick={() => { history.goBack() }} className="reset-btn">
                                                Exit
                                            </Link>
                                        </div>
                                    </div>
                                    <div className="mt-5">
                                        <table className="table table-bordered mb-0">
                                            <tbody>
                                                <tr>
                                                    <td
                                                        className="text-start font"
                                                        style={{ width: "40%", padding: "4px 8px", lineHeight: "1.2" }}
                                                    >
                                                        Format No. - SonaBLW/F/QA/29
                                                    </td>

                                                    <td
                                                        className="text-center font"
                                                        style={{ width: "20%", padding: "4px 8px", lineHeight: "1.2" }}
                                                    >
                                                        Rev. No. - 03
                                                    </td>

                                                    <td
                                                        className="text-end font"
                                                        style={{ width: "40%", padding: "4px 8px", lineHeight: "1.2" }}
                                                    >
                                                        w.e.f. - 17.02.2024
                                                    </td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div >
        )
    }
}