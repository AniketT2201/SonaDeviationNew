

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
            if (ApprovalMatrixdata) {
                displayWorkflow();
            }
        }, [ApprovalMatrixdata])

        const displayWorkflow = () => {
            const wf: JSX.Element[] = [];

            // const _wf = approverJson.filter((item) => item.required === true);
            let isActive;
            let notActive = false;
            ApprovalMatrixdata.filter(m => m.required === true).forEach((m, i) => {
                //if (m.required === true) {
                if (notActive === false && Stage !== 99) {
                    if (Stage === i) {
                        isActive = 'activeApprover';
                        notActive = true;
                    }
                    else {
                        isActive = 'beforeactiveApprover';
                    }
                }
                else {
                    isActive = 'overrideStage';
                }

                wf.push(
                    <ul className="main-menu">
                        <li className={`${m.Role} ${isActive}`.trim()}>
                            {m.User}
                        </li>
                    </ul>
                );
                //}
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
                        "PlantHeadDate", "PlantHeadEffectiveness", "AttachmentFiles", "Comments" 
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

        return (
            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
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
                                                    <a href={getFullUrl(DeviationAttachments[0].ServerRelativeUrl)} target="_blank" rel="noopener noreferrer" >
                                                        {(DeviationAttachments[0].FileName)}
                                                    </a>
                                                ) : (
                                                    <span>No attachment found.</span>
                                                )}
                                            </div>
                                        </div>
                                        {isEngineeringDashboard && (
                                            <div className='col-md-4'>
                                                <label htmlFor="Comments" className='font'>Comments </label> : &nbsp;&nbsp;
                                                <div><label className='fonttext'>{Comments}</label></div>
                                            </div>
                                        )}
                                    </div>
                                </div>

                                {/* 🔹 Table */}
                                <div className='RequestTable'>
                                    <div style={{ overflowX: "auto", width: "100%" }}>
                                        <table>
                                            <colgroup>
                                                <col style={{ width: "60px" }} />   {/* S.N */}
                                                <col style={{ width: "220px" }} />  {/* Parameters */}
                                                <col style={{ width: "220px" }} />  {/* Sub Parameters */}
                                                {showOthersColumn && <col style={{ width: "220px"}} />}  {/* Others */}
                                                <col style={{ width: "180px" }} />
                                                <col style={{ width: "180px" }} />
                                                <col style={{ width: "150px" }} />
                                                <col style={{ width: "100px" }} />
                                                <col style={{ width: "250px" }} />
                                                <col style={{ width: "250px" }} />
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
                                        <div className='text-center'>
                                            <Link to="#" onClick={() => { history.goBack() }} className="reset-btn">
                                                Exit
                                            </Link>
                                        </div>
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