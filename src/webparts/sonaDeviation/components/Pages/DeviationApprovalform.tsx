

import * as React from 'react';
import { useState, useEffect } from 'react';
import type { ISonaDeviationProps } from '../ISonaDeviationProps';
import { TextField, DefaultButton, Dropdown, IDropdownOption } from '@fluentui/react';
import { Link, useHistory, useLocation } from "react-router-dom";
import logo from "../../assets/sona-comstarlogo.png";
import 'bootstrap/dist/css/bootstrap.min.css';
import "../Pages/Css/NewRequest.scss";

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import { useParams } from "react-router-dom";

import UserProfileOps from '../../service/BAL/UserProfile';
import { IUserProfile } from '../../service/INTERFACE/IUserProfile';
import { set } from '@microsoft/sp-lodash-subset/lib/index';

interface IRouteParams {
    id: string;
}

export const DeviationApproverForm = (props: ISonaDeviationProps) => {
    {
        const { id } = useParams<IRouteParams>();
        const history = useHistory();
        const location = useLocation<any>();

        const [Details, setDetails] = React.useState<any[]>([]);
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

        const [QualityApprover, setQualityApprover] = React.useState(false);
        const [EngineeringApprover, setEngineeringApprover] = React.useState(false);
        const [COOApprover, setCOOApprover] = React.useState(false);
        const [PlantHeadApprover, setPlantHeadApprover] = React.useState(false);

        const [Remark, setRemark] = React.useState("");
        const [Functional, setFunctional] = React.useState("");
        const [Fitment, setFitment] = React.useState("");
        const [SpecialNote, setSpecialNote] = React.useState("");
        const [SelectedDeviationRecommended, setSelectedDeviationRecommended] = React.useState<number | undefined>(undefined);
        // const [DeviationRecommended, setDeviationRecommended] = React.useState("");
        const [Name, setName] = React.useState("");
        const [QualityProductionDate, setQualityProductionDate] = React.useState("");
        const [ProcessingDate, setProcessingDate] = React.useState("");

        const [EngineeringRemark, setEngineeringRemark] = React.useState("");
        const [EngineeringFunctional, setEngineeringFunctional] = React.useState("");
        const [EngineeringFitment, setEngineeringFitment] = React.useState("");
        const [EngineeringSpecialNote, setEngineeringSpecialNote] = React.useState("");
        const [SelectedEngineeringDeviationRecommended, setSelectedEngineeringDeviationRecommended] = useState<number | undefined>(undefined);
        const [EngineeringName, setEngineeringName] = React.useState("");
        const [EngineeringProductionDate, setEngineeringProductionDate] = React.useState("");
        const [EngineeringDate, setEngineeringDate] = React.useState("");
        const [SelectedSpecialApp, setSelectedSpecialApp] = React.useState<string>("No");

        const [COORemark, setCOORemark] = React.useState("");
        const [RegionComment, setRegionComment] = React.useState("");

        const [PlantHeadRemark, setPlantHeadRemark] = React.useState("");
        const [SelectedPlantHeadDeviationAccepted, setSelectedPlantHeadDeviationAccepted] = useState<number | undefined>(undefined);
        const [PlantHeadProductionDate, setPlantHeadProductionDate] = React.useState("");
        const [PlantHeadDate, setPlantHeadDate] = React.useState("");
        const [PlantHeadEffectiveness, setPlantHeadEffectiveness] = React.useState("");

        const [DeviationAttachments, setDeviationAttachments] = React.useState<any[]>([]);

        const [NextsaApprover, setNextsaApprover] = React.useState(false);

        const [WorkflowJSX, setWorkflowJSX] = React.useState<JSX.Element[]>([]);

        const [showPopup, setShowPopup] = useState(false);
        const [actionType, setActionType] = useState<"rework" | "reject" | "">("");
        const textareaRef = React.useRef<any>(null);
        const [Comments, setComments] = React.useState("");
        const [selectedRiskAssessment, setSelectedRiskAssessment] = React.useState("");
        const [RiskAssessmentDescription, setRiskAssessmentDescription] = React.useState("");

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

        const yesNoOptions: IDropdownOption[] = [
            { key: 1, text: "Yes" },
            { key: 2, text: "No" },
        ];

        const YesNoOptions: IDropdownOption[] = [
            { key: "Yes", text: "Yes" },
            { key: "No", text: "No" },
        ];

        React.useEffect(() => {
            sp.setup({
                spfxContext: props.currentSPContext
            });

            if (id) {
                loadItem(Number(id));
            }

        }, [id]);

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
                wf.push(
                    <ul className="main-menu">
                        <li className={`${m.Role} ${isActive}`.trim()}>
                            {m.User}
                        </li>
                    </ul>
                );
            });

            setWorkflowJSX(wf);
        };

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
                        "EngineeringName", "EngineeringProductionDate", "EngineeringDate", "SpecialApproval",
                        "PlantHeadRemark", "PlantHeadDeviationAccepted", "PlantHeadProductionDate", "COORemark",
                        "PlantHeadDate", "PlantHeadEffectiveness", "NextApprover/ID", "AttachmentFiles", "Comments", "RiskAssessment", "RiskAssessmentDescription"
                    ).expand("NextApprover,AttachmentFiles")();
                setPartNo(item.PartNo || "");
                setDeviationNo(item.DeviationNo || "");
                setConcernedDepartment(item.SupplierName || "");
                setBatchNo(item.BatchNo || "");
                setPlantName(item.PlantName || "");
                setProductionQTY(item.ProductionQTY || "");
                setRequestorName(item.RequestorName || "");
                setProductionDate(item.ProductionDate || "");

                setStage(item.Stage || "");
                setDeviationAttachments(item.AttachmentFiles || []);

                setRemark(item.Remark || "");
                setFunctional(item.Functional || "");
                setFitment(item.Fitment || "");
                setSpecialNote(item.SpecialNote || "");
                if (item.DeviationRecommended === "Yes") {
                    setSelectedDeviationRecommended(1);
                } else if (item.DeviationRecommended === "No") {
                    setSelectedDeviationRecommended(0);
                } else {
                    setSelectedDeviationRecommended(undefined);
                }
                setName(item.QualityName || "");
                setQualityProductionDate(item.QualityProductionDate || "");
                setProcessingDate(item.QualityDate || "");

                setEngineeringRemark(item.EngineeringRemark || "");
                setEngineeringFunctional(item.EngineeringFunctional || "");
                setEngineeringFitment(item.EngineeringFitment || "");
                setEngineeringSpecialNote(item.EngineeringSpecialNote || "");
                if (item.EngineeringDeviationRecommended === "Yes") {
                    setSelectedEngineeringDeviationRecommended(1);
                } else if (item.EngineeringDeviationRecommended === "No") {
                    setSelectedEngineeringDeviationRecommended(0);
                } else {
                    setSelectedEngineeringDeviationRecommended(undefined);
                }
                setEngineeringName(item.EngineeringName || "");
                setEngineeringProductionDate(item.EngineeringProductionDate || "");
                setEngineeringDate(item.EngineeringDate || "");
                setSelectedSpecialApp(item.SpecialApproval || "No");

                setCOORemark(item.COORemark || "");

                setPlantHeadRemark(item.PlantHeadRemark || "");
                setSelectedPlantHeadDeviationAccepted(item.SelectedPlantHeadDeviationAccepted || "");
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
                // for (var i = 0; i < RoletypeArr.length; i++) {
                //     let stage = await filterdata(item, RoletypeArr[i]);
                //     if (stage === true) {
                //         if (RoletypeArr[i] === "Engineering") { setEngineeringApprover(true) };
                //         if (RoletypeArr[i] === "COO") { setCOOApprover(true); };
                //         if (RoletypeArr[i] === "Quality") { setQualityApprover(true); };
                //         if (RoletypeArr[i] === "PlantHead") { setPlantHeadApprover(true); };
                //     }
                // };


                for (var i = 0; i < RoletypeArr.length; i++) {
                    let stage = await filterdata(item, RoletypeArr[i]);

                    if (stage === true) {
                        if (RoletypeArr[i] === "Engineering") { setEngineeringApprover(true) }
                        if (RoletypeArr[i] === "COO") { setCOOApprover(true) }
                        if (RoletypeArr[i] === "Quality") { setQualityApprover(true) }
                        if (RoletypeArr[i] === "PlantHead") { setPlantHeadApprover(true) }
                    }
                }


                if (item.NextApprover?.ID === props.id) {
                    setNextsaApprover(true);
                }

            } catch (error) {
                console.error("Error loading BG item:", error);
            }
        };

        const ApproveRequest = async () => {
            if (isSubmitting) return;

            setIsSubmitting(true);

            try {

                const itemId = Number(id);
                const currentUser = await sp.web.currentUser();

                // =====================================
                // PARSE APPROVAL MATRIX
                // =====================================

                let nextapproverdata;
                let uptstatus;
                let updateFields;
                if (ApprovalMatrixdata.filter(m => m.required === true).length === Stage + 1) {
                    nextapproverdata = null;
                    uptstatus = "Approved";
                }
                else {
                    nextapproverdata = ApprovalMatrixdata.filter(m => m.Stage === Stage + 1 && m.required === true);
                    uptstatus = "Pending for Approval";
                }


                let CurrentApprover = ApprovalMatrixdata.filter(m => m.Stage === Stage && m.required === true);
                const nextIndex = Stage + 1;
                let nextApprover;
                let currentstatus;
                if (nextapproverdata !== null && nextapproverdata !== undefined && nextapproverdata !== '') {
                    nextApprover = nextapproverdata[0];
                    currentstatus = `Submitted to ${nextapproverdata[0].Role}`
                }
                const now = new Date();
                const formattedDate = now.toLocaleDateString("en-GB");
                //wf             
                let workflowHistory = [...WorkflowHistorydata];
                const newHistoryEntry = {
                    CurrentApprover: currentUser.Title,
                    ActionTaken: `${CurrentApprover[0].Role} Approved`,
                    Date: formattedDate,
                    CurrentStatus: currentstatus
                };
                workflowHistory.push(newHistoryEntry);
                const updatedHistory = JSON.stringify(workflowHistory);
                //           
                const Role = ApprovalMatrixdata.filter(m => m.Stage === Stage && m.required === true)[0].Role;
                if (Role === 'Engineering') {

                    // ✅ VALIDATION
                    // if (!SelectedSpecialApp) {
                    //     alert("Please select Special Approval");
                    //     setIsSubmitting(false);
                    //     return;
                    // }

                    if (SelectedEngineeringDeviationRecommended === 2) {
                        alert("Deviation Recommended is No. You cannot approve this request.");
                        setIsSubmitting(false);
                        return;
                    }
                    if (!SelectedEngineeringDeviationRecommended) {
                        alert("Please select Engineering Deviation Recommended.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!EngineeringRemark) {
                        alert("Please select Engineering Remark.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!EngineeringFunctional) {
                        alert("Please select Engineering Functional.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!EngineeringFitment) {
                        alert("Please select Engineering Fitment.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!EngineeringSpecialNote) {
                        alert("Please select Engineering Special Note.");
                        setIsSubmitting(false);
                        return;
                    }
                    const specialApprovalValue = SelectedSpecialApp || "No";
                    console.log("SelectedSpecialApp:", SelectedSpecialApp);
                    if (specialApprovalValue === "No") {
                        updateFields = {
                            Stage: nextIndex || null,
                            Status: uptstatus,
                            WorkflowHistory: updatedHistory,
                            EngineeringRemark: EngineeringRemark || null,
                            EngineeringFunctional: EngineeringFunctional || null,
                            EngineeringFitment: EngineeringFitment || null,
                            EngineeringSpecialNote: EngineeringSpecialNote || null,
                            EngineeringDeviationRecommended: SelectedEngineeringDeviationRecommended === 1 ? "Yes" : SelectedEngineeringDeviationRecommended === 2 ? "No" : null,
                            EngineeringName: EngineeringName || null,
                            EngineeringProductionDate: EngineeringProductionDate ? new Date(EngineeringProductionDate) : null,
                            EngineeringDate: EngineeringDate ? new Date(EngineeringDate) : null,
                            NextApproverId: nextApprover
                                ? parseInt(nextApprover.UserID)
                                : null,
                            SpecialApproval: SelectedSpecialApp || null
                        }
                    }
                    else if (specialApprovalValue === "Yes") {
                        let aspercoo: any = [];
                        let newAppdata = ApprovalMatrixdata.filter(m => m.Role === "COO");
                        let oldcoowfHis = [...WorkflowHistorydata];
                        const COOHistoryEntry = {
                            CurrentApprover: currentUser.Title,
                            ActionTaken: `${CurrentApprover[0].Role} Approved`,
                            Date: formattedDate,
                            CurrentStatus: `Submitted to ${newAppdata[0].Role}`
                        };
                        oldcoowfHis.push(COOHistoryEntry);
                        let RoletypeArr = 
                            SelectedSpecialApp === "Yes"
                            ? ["Initiator", "Engineering", "COO"]
                            : ["Initiator", "Engineering", "Quality", "PlantHead"];
                        for (var j = 0; j < RoletypeArr.length; j++) {
                            let data = ApprovalMatrixdata.filter(m => m.Role === RoletypeArr[j]);
                            aspercoo.push({
                                Role: data[0]?.Role,
                                User: data[0]?.User,
                                UserID: data[0]?.UserID,
                                UserEmail: data[0]?.UserEmail,
                                PendingText: ('Level ' + j).toString(),
                                Stage: j,
                                required: true
                            });

                        }

                        updateFields = {
                            Stage: nextIndex || null,
                            Status: uptstatus,
                            WorkflowHistory: oldcoowfHis ? JSON.stringify(oldcoowfHis) : null,
                            EngineeringRemark: EngineeringRemark || null,
                            EngineeringFunctional: EngineeringFunctional || null,
                            EngineeringFitment: EngineeringFitment || null,
                            EngineeringSpecialNote: EngineeringSpecialNote || null,
                            EngineeringDeviationRecommended: SelectedEngineeringDeviationRecommended === 1 ? "Yes" : SelectedEngineeringDeviationRecommended === 2 ? "No" : null,
                            EngineeringName: EngineeringName || null,
                            EngineeringProductionDate: EngineeringProductionDate ? new Date(EngineeringProductionDate) : null,
                            EngineeringDate: EngineeringDate ? new Date(EngineeringDate) : null,
                            NextApproverId: newAppdata[0]
                                ? parseInt(newAppdata[0].UserID)
                                : null,
                            SpecialApproval: SelectedSpecialApp || null,
                            ApprovalMatrix: JSON.stringify(aspercoo)
                        }
                    }
                }
                else if (Role === 'COO') {

                    if (!COORemark) {
                        alert("Please select COO Remark.");
                        setIsSubmitting(false);
                        return;
                    }

                    const finalStatus = "Approved";
                    const finalStage = 0;

                    // Update workflow history for final approval
                    let workflowHistory = [...WorkflowHistorydata];
                    const finalHistoryEntry = {
                        CurrentApprover: currentUser.Title,
                        ActionTaken: `${CurrentApprover[0].Role} Approved`,
                        Date: formattedDate,
                        CurrentStatus: "Approved"
                    };
                    workflowHistory.push(finalHistoryEntry);

                    updateFields = {
                        Stage: finalStage,
                        Status: finalStatus,
                        COORemark: COORemark || null,
                        WorkflowHistory: JSON.stringify(workflowHistory),
                        NextApproverId: null
                    };
                }
                //---------------------   this is for COO Approve  ----------------------------------

                // else if (Role === 'COO') {
                //     updateFields = {
                //         Stage: nextIndex || null,
                //         Status: uptstatus,
                //         COORemark: COORemark || null,
                //         WorkflowHistory: updatedHistory,
                //         NextApproverId: nextApprover
                //             ? parseInt(nextApprover.UserID)
                //             : null
                //     };
                // }

                // ----------------------------------------------------------------------------------

                else if (Role === 'Quality') {

                    if (SelectedDeviationRecommended === 2) {
                        alert("Deviation Recommended is No. You cannot approve this request.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!SelectedDeviationRecommended) {
                        alert("Please select Quality Deviation Recommended.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!Remark) {
                        alert("Please select Quality Remark.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!Functional) {
                        alert("Please select Quality Functional.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!Fitment) {
                        alert("Please select Quality Fitment.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!SpecialNote) {
                        alert("Please select Quality Special Note.");
                        setIsSubmitting(false);
                        return;
                    }


                    updateFields = {
                        Stage: nextIndex || null,
                        WorkflowHistory: updatedHistory,
                        Remark: Remark || null,
                        Functional: Functional || null,
                        Fitment: Fitment || null,
                        SpecialNote: SpecialNote || null,
                        DeviationRecommended: SelectedDeviationRecommended === 1 ? "Yes" : SelectedDeviationRecommended === 2 ? "No" : null,
                        QualityName: Name || null,
                        QualityProductionDate: QualityProductionDate ? new Date(QualityProductionDate) : null,
                        QualityDate: ProcessingDate ? new Date(ProcessingDate) : null,
                        Status: uptstatus,
                        NextApproverId: nextApprover
                            ? parseInt(nextApprover.UserID)
                            : null
                    }
                }
                else if (Role === 'PlantHead') {

                    if (!PlantHeadRemark) {
                        alert("Please select Plant Head Remark.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!SelectedPlantHeadDeviationAccepted) {
                        alert("Please select Plant Head Deviation Accepted.");
                        setIsSubmitting(false);
                        return;
                    }

                    if (!PlantHeadEffectiveness) {
                        alert("Please select Plant Head Effectiveness.");
                        setIsSubmitting(false);
                        return;
                    }

                    updateFields = {
                        Stage: nextIndex || null,
                        Status: uptstatus,
                        WorkflowHistory: updatedHistory,
                        PlantHeadRemark: PlantHeadRemark || null,
                        PlantHeadDeviationAccepted: SelectedPlantHeadDeviationAccepted === 1 ? "Yes" : SelectedPlantHeadDeviationAccepted === 2 ? "No" : null,
                        PlantHeadProductionDate: PlantHeadProductionDate ? new Date(PlantHeadProductionDate) : null,
                        PlantHeadDate: PlantHeadDate ? new Date(PlantHeadDate) : null,
                        PlantHeadEffectiveness: PlantHeadEffectiveness || null,
                        NextApproverId: nextApprover
                            ? parseInt(nextApprover.UserID)
                            : null
                    }
                }


                // =====================================
                // UPDATE ITEM
                // =====================================

                await sp.web.lists
                    .getByTitle("DeviationDetails")
                    .items.getById(itemId)
                    .update(updateFields);

                alert(`Request successfully approved by ${CurrentApprover[0].Role}`);
                history.push(`/${CurrentApprover[0].Role}Dashboard`);
            }
            catch (error) {
                console.error(error);
                alert("Approval failed");
            }
            finally {
                setIsSubmitting(false);
            }
        }

        const handleRework = async () => {

            if (isSubmitting) return;
            setIsSubmitting(true);

            try {
                const itemIdNum = Number(id);
                const currentUser = await sp.web.currentUser();

                let CurrentApprover = ApprovalMatrixdata.filter(
                    m => m.Stage === Stage && m.required === true
                );

                const role = CurrentApprover[0].Role;

                // 🔁 Find previous stage
                const prevStageIndex = Stage - 1;

                if (prevStageIndex < 0) {
                    alert("Cannot send back. Already at initial stage.");
                    setIsSubmitting(false);
                    return;
                }

                const prevApproverData = ApprovalMatrixdata.filter(
                    m => m.Stage === prevStageIndex && m.required === true
                );

                const prevApprover = prevApproverData[0];

                const now = new Date();
                const formattedDate = now.toLocaleDateString("en-GB");

                // 📜 Workflow History
                let workflowHistory = [...WorkflowHistorydata];

                const reworkEntry = {
                    CurrentApprover: currentUser.Title,
                    Comment: RegionComment,
                    ActionTaken: `${role} Send Back`,
                    Date: formattedDate,
                    CurrentStatus: `Send Back To ${prevApprover.Role}`
                };

                workflowHistory.push(reworkEntry);

                const updatedHistory = JSON.stringify(workflowHistory);

                // 🧾 Update fields
                let updateFields: any = {
                    Status: "Send Back",
                    Stage: 0,
                    NextApproverId: prevApprover
                        ? parseInt(prevApprover.UserID)
                        : null,
                    WorkflowHistory: updatedHistory
                };

                // 📌 Save remarks based on role
                if (role === "Quality") {
                    updateFields.Remark = Remark || null;
                }
                else if (role === "Engineering") {
                    updateFields.EngineeringRemark = EngineeringRemark || null;
                }
                else if (role === "COO") {
                    updateFields.COORemark = COORemark || null;
                }
                else if (role === "PlantHead") {
                    updateFields.PlantHeadRemark = PlantHeadRemark || null;
                }

                await sp.web.lists
                    .getByTitle("DeviationDetails")
                    .items.getById(itemIdNum)
                    .update(updateFields);

                alert(`Request sent back by ${role}`);
                history.push(`/${role}Dashboard`);
            }
            catch (error) {
                console.error(error);
                alert("Send back failed");
            }
            finally {
                setIsSubmitting(false);
            }
        };

        const handleReject = async () => {

            if (isSubmitting) return;
            setIsSubmitting(true);

            try {
                const itemIdNum = Number(id);
                const currentUser = await sp.web.currentUser();

                let CurrentApprover = ApprovalMatrixdata.filter(
                    m => m.Stage === Stage && m.required === true
                );

                const now = new Date();
                const formattedDate = now.toLocaleDateString("en-GB");

                let workflowHistory = [...WorkflowHistorydata];

                const rejectEntry = {
                    CurrentApprover: currentUser.Title,
                    Comment: RegionComment,
                    ActionTaken: `${CurrentApprover[0].Role} Rejected`,
                    Date: formattedDate,
                    CurrentStatus: "Rejected"
                };

                workflowHistory.push(rejectEntry);

                const updatedHistory = JSON.stringify(workflowHistory);

                const updateFields = {
                    Status: "Rejected",
                    NextApproverId: null,
                    WorkflowHistory: updatedHistory,
                    Stage: null
                };

                await sp.web.lists
                    .getByTitle("DeviationDetails")
                    .items.getById(itemIdNum)
                    .update(updateFields);

                alert(`Request rejected by ${CurrentApprover[0].Role}`);
                history.push(`/${CurrentApprover[0].Role}Dashboard`);
            }
            catch (error) {
                console.error(error);
                alert("Rejection failed");
            }
            finally {
                setIsSubmitting(false);
            }
        };

        const getFullUrl = (relativeUrl: string) => {
            const baseUrl = props.currentSPContext.pageContext.web.absoluteUrl.split('/sites')[0];
            return baseUrl + relativeUrl;
        };

        const AddComment = async () => {
            if (!RegionComment || RegionComment.trim() === "") {
                alert("Please enter a comment before submitting.");
                return;
            }

            try {
                if (actionType === "rework") {
                    await handleRework();
                } else if (actionType === "reject") {
                    await handleReject();
                }

                setShowPopup(false);
                setRegionComment("");
                setActionType("");

            } catch (error) {
                console.error("Error:", error);
            }
        };

        const handleClosePopup = () => {
            setShowPopup(false);
            setRegionComment(""); // ✅ reset textarea value
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
                                <h1>REQUEST FOR DEVIATION </h1>
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
                                            <label className='fonttext'>{ProductionDate}</label>
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
                                            <div className='col-md-3'>
                                                <label htmlFor="Comments" className='font'>Comments </label> <br></br>
                                                <div><label className='fonttext' style={{ fontWeight: 'bold', color: 'red'}}>{Comments}</label></div>
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

                                {(EngineeringApprover) && (
                                    <div>
                                        <div className='heading1'>
                                            <label>Recommendation Engineering</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Special Approval <span className="Mantorystar">* If Special Approval is set to Yes, the request will go to the COO for approval.</span></label>
                                                    <Dropdown
                                                        className='formtext-control'
                                                        options={YesNoOptions}
                                                        selectedKey={SelectedSpecialApp}
                                                        onChange={(e, option) => {
                                                            console.log(option?.key);
                                                            setSelectedSpecialApp(option?.key as string)
                                                        }
                                                        }
                                                    />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Deviation Recommended <span className="Mantorystar">*</span></label>
                                                    <Dropdown
                                                        className='formtext-control'
                                                        options={yesNoOptions}
                                                        selectedKey={SelectedEngineeringDeviationRecommended}
                                                        onChange={(e, option) =>
                                                            setSelectedEngineeringDeviationRecommended(option?.key as number)
                                                        }
                                                    />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Functional <span className="Mantorystar">*</span></label>
                                                    <input type="text"
                                                        value={EngineeringFunctional}

                                                        className='form-control'
                                                        onChange={e => setEngineeringFunctional(e.target.value)} />
                                                </div>
                                            </div>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Remarks <span className="Mantorystar">*</span></label>
                                                    <textarea className="form-control" value={EngineeringRemark}
                                                        onChange={(e) => {
                                                            const textarea = e.currentTarget;
                                                            textarea.style.height = "auto";              // reset height
                                                            textarea.style.height = textarea.scrollHeight + "px"; // expand
                                                            setEngineeringRemark(textarea.value);
                                                        }}
                                                        style={{ resize: "none", overflowY: "hidden" }}
                                                    />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Fitment <span className="Mantorystar">*</span></label>
                                                    <textarea className="form-control"
                                                        value={EngineeringFitment}
                                                        onChange={(e) => {
                                                            const textarea = e.currentTarget;
                                                            textarea.style.height = "auto";              // reset height
                                                            textarea.style.height = textarea.scrollHeight + "px"; // expand
                                                            setEngineeringFitment(textarea.value);
                                                        }}
                                                        style={{ resize: "none", overflowY: "hidden" }}
                                                    />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Special Note <span className="Mantorystar">*</span></label>
                                                    <textarea className="form-control"
                                                        value={EngineeringSpecialNote}
                                                        onChange={(e) => {
                                                            const textarea = e.currentTarget;
                                                            textarea.style.height = "auto";              // reset height
                                                            textarea.style.height = textarea.scrollHeight + "px"; // expand
                                                            setEngineeringSpecialNote(textarea.value);
                                                        }}
                                                        style={{ resize: "none", overflowY: "hidden" }}
                                                    />
                                                </div>
                                            </div>
                                            {/* <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Name </label>
                                                    <input type="text"
                                                        value={EngineeringName}
                                                        className='form-control'
                                                        onChange={e => setEngineeringName(e.target.value)} />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Production Date </label>
                                                    <input type="date" value={EngineeringProductionDate} className='form-control' onChange={e => setEngineeringProductionDate(e.target.value)} />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Date </label>
                                                    <input type="date" value={EngineeringDate} className='form-control' onChange={e => setEngineeringDate(e.target.value)} />
                                                </div>
                                            </div> */}
                                        </div>
                                    </div>
                                )}
                                {(COOApprover || QualityApprover || PlantHeadApprover) && (
                                    <div>
                                        <div className='heading1'>
                                            <label>Recommendation Engineering</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'> Special Approval <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{SelectedSpecialApp}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Deviation Recommended <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{SelectedEngineeringDeviationRecommended}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Functional <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{EngineeringFunctional}</label>
                                                </div>
                                            </div>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Remarks <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{EngineeringRemark}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Fitment <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{EngineeringFitment}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Special Note <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
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
                                    </div>
                                )}
                                {COOApprover && SelectedSpecialApp === "Yes" && (
                                    <div>
                                        <div className='heading1'>
                                            <label>Recommendation COO</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Remarks <span className="Mantorystar">*</span></label>
                                                    <textarea className="form-control" value={COORemark} onChange={e => setCOORemark(e.target.value)}></textarea>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                )}
                                {(QualityApprover || PlantHeadApprover) && SelectedSpecialApp === "Yes" && (
                                    <div>
                                        <div className='heading1'>
                                            <label>Recommendation COO</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Remarks <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{COORemark}</label>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                )}
                                {QualityApprover && SelectedSpecialApp === "No" && (
                                    <div>
                                        <div className='heading1'>
                                            <label>Recommendation Quality</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Deviation Recommended <span className="Mantorystar">*</span></label>
                                                    <Dropdown
                                                        className='formtext-control'
                                                        options={yesNoOptions}
                                                        selectedKey={SelectedDeviationRecommended}
                                                        onChange={(e, option) =>
                                                            setSelectedDeviationRecommended(option?.key as number)
                                                        }
                                                    />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Functional <span className="Mantorystar">*</span></label>
                                                    <input type="text" value={Functional} className='form-control' onChange={e => setFunctional(e.target.value)} />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Remarks <span className="Mantorystar">*</span></label>
                                                    <textarea className="form-control" value={Remark}
                                                        onChange={(e) => {
                                                            const textarea = e.currentTarget;
                                                            textarea.style.height = "auto";
                                                            textarea.style.height = textarea.scrollHeight + "px";
                                                            setRemark(textarea.value);
                                                        }}
                                                        style={{ resize: "none", overflowY: "hidden" }}
                                                    />
                                                </div>

                                            </div>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Fitment <span className="Mantorystar">*</span></label>
                                                    <textarea className="form-control" value={Fitment}
                                                        onChange={(e) => {
                                                            const textarea = e.currentTarget;
                                                            textarea.style.height = "auto";
                                                            textarea.style.height = textarea.scrollHeight + "px";
                                                            setFitment(textarea.value);
                                                        }}
                                                        style={{ resize: "none", overflowY: "hidden" }}
                                                    />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Special Note <span className="Mantorystar">*</span></label>
                                                    <textarea className="form-control" value={SpecialNote}
                                                        onChange={(e) => {
                                                            const textarea = e.currentTarget;
                                                            textarea.style.height = "auto";
                                                            textarea.style.height = textarea.scrollHeight + "px";
                                                            setSpecialNote(textarea.value);
                                                        }}
                                                        style={{ resize: "none", overflowY: "hidden" }}
                                                    />
                                                </div>
                                            </div>
                                        </div>
                                        {/* <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Name </label>
                                                    <input type="text" value={Name} className='form-control' onChange={e => setName(e.target.value)} />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Production Date </label>
                                                    <input type="date" value={QualityProductionDate} className='form-control' onChange={e => setQualityProductionDate(e.target.value)} />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Date </label>
                                                    <input type="date" value={ProcessingDate} className='form-control' onChange={e => setProcessingDate(e.target.value)} />
                                                </div>
                                            </div> */}
                                    </div>
                                )}
                                {PlantHeadApprover && SelectedSpecialApp === "No" && (
                                    <div>
                                        <div className='heading1'>
                                            <label>Recommendation Quality</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Deviation Recommended <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{SelectedDeviationRecommended}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Functional <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{Functional}</label>
                                                </div>
                                                <div className='col-md-4 '>
                                                    <label className='font'>Remarks <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{Remark}</label>
                                                </div>
                                            </div>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Fitment <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{Fitment}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Special Note <span className="Mantorystar">*</span></label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{SpecialNote}</label>
                                                </div>
                                            </div>
                                            {/* <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Name </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{Name}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Production Date </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{QualityProductionDate}</label>
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Date </label> : &nbsp;&nbsp;
                                                    <label className='fonttext'>{formatDate(ProcessingDate)}</label>
                                                </div>
                                            </div> */}
                                        </div>
                                    </div>
                                )}
                                {PlantHeadApprover && SelectedSpecialApp === "No" && (
                                    <div>
                                        <div className='heading1'>
                                            <label>Recommendation PlantHead</label>
                                        </div>
                                        <div className='main-formcontainer'>
                                            <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Deviation Accepted <span className="Mantorystar">*</span></label>
                                                    <Dropdown
                                                        className='formtext-control'
                                                        options={yesNoOptions}
                                                        selectedKey={SelectedPlantHeadDeviationAccepted}
                                                        onChange={(e, option) =>
                                                            setSelectedPlantHeadDeviationAccepted(option?.key as number)
                                                        }
                                                    />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Remarks <span className="Mantorystar">*</span></label>
                                                    <textarea className="form-control" value={PlantHeadRemark}
                                                        onChange={(e) => {
                                                            const textarea = e.currentTarget;
                                                            textarea.style.height = "auto";
                                                            textarea.style.height = textarea.scrollHeight + "px";
                                                            setPlantHeadRemark(textarea.value);
                                                        }}
                                                        style={{ resize: "none", overflowY: "hidden" }}
                                                    />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Effectiveness <span className="Mantorystar">*</span></label>
                                                    <textarea className="form-control" value={PlantHeadEffectiveness}
                                                        onChange={(e) => {
                                                            const textarea = e.currentTarget;
                                                            textarea.style.height = "auto";
                                                            textarea.style.height = textarea.scrollHeight + "px";
                                                            setPlantHeadEffectiveness(textarea.value);
                                                        }}
                                                        style={{ resize: "none", overflowY: "hidden" }}
                                                    />
                                                </div>
                                            </div>
                                            {/* <div className='row mb-20'>
                                                <div className='col-md-4'>
                                                    <label className='font'>Production Date </label>
                                                    <input type="date" value={PlantHeadProductionDate} className='form-control' onChange={e => setPlantHeadProductionDate(e.target.value)} />
                                                </div>
                                                <div className='col-md-4'>
                                                    <label className='font'>Date </label>
                                                    <input type="date" value={PlantHeadDate} className='form-control' onChange={e => setPlantHeadDate(e.target.value)} />
                                                </div>
                                            </div> */}
                                        </div>
                                    </div>
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
                                        <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                                            {NextsaApprover && (
                                                <>
                                                    < a onClick={() => !isSubmitting && ApproveRequest()} className="submit-btn">
                                                        Accept
                                                    </a>
                                                    <a onClick={() => { if (isSubmitting) return; setActionType("rework"); setShowPopup(true); }} className="Rework-btn">
                                                        Rework
                                                    </a>
                                                    <a onClick={() => { if (isSubmitting) return; setActionType("reject"); setShowPopup(true); }} className="Reject-btn">
                                                        Reject
                                                    </a>
                                                </>
                                            )}
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
                {showPopup && (
                    <div className="popup-overlay">
                        <div className="popup-contentone">
                            <div className="blockpopup">
                                <div className="block-headpopup">
                                    <div className="head-leftpopup">
                                        <span>Comment &nbsp;<span style={{ color: "red" }}> * </span></span>
                                    </div>
                                    <div className="popupclose">
                                        <button className="close-btnpopup" onClick={handleClosePopup}>X</button>
                                    </div>
                                </div>

                                <div className="block-contentpopup">
                                    <div className="Main-popup">
                                        <div className="container">
                                            <div className="row">
                                                <div className="col-md-12">
                                                    <div className="Main-ContentVendorPopup" style={{ paddingTop: "1rem" }}>
                                                        <div className='row mb-20'>
                                                            <div className="col-md-12">
                                                                <textarea
                                                                    ref={textareaRef}
                                                                    className="form-control"
                                                                    value={RegionComment}
                                                                    onChange={(e) => {
                                                                        const textarea = e.currentTarget;
                                                                        textarea.style.height = "auto";
                                                                        textarea.style.height = textarea.scrollHeight + "px";
                                                                        setRegionComment(textarea.value);
                                                                    }}
                                                                    style={{ resize: "none", overflowY: "hidden", height: "auto" }}
                                                                />
                                                            </div>
                                                        </div>
                                                        <div className='row' style={{ paddingTop: "5px" }}>
                                                            <div className='col-md-12'>
                                                                <div className='text-center'>
                                                                    <a onClick={() => AddComment()} className="submit-btn">
                                                                        Submit
                                                                    </a>
                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                )}
            </div >
        )
    }
}