import * as React from 'react';
import { useState, useRef } from 'react';
import type { ISonaDeviationProps } from '../ISonaDeviationProps';
import { TextField, DefaultButton, Dropdown, IDropdownOption } from '@fluentui/react';
import { Label } from '@fluentui/react/lib/Label';
import logo from "../../assets/sona-comstarlogo.png";
import "@pnp/sp/folders";

import { Link } from "react-router-dom";
import "./Css/NewRequest.scss"

import { sp } from "@pnp/sp/presets/all";

import { Web } from "@pnp/sp/webs";

import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

import 'bootstrap/dist/css/bootstrap.min.css';

import SPCRUDOPS from "../../service/BAL/spcrud";

import { Upload, Button } from "antd";
import { UploadOutlined } from "@ant-design/icons";

import { useParams } from "react-router-dom";

import { useHistory } from 'react-router-dom';


import ParameterTypeOps from '../../service/BAL/Parameter';
import { IParameterType } from '../../service/INTERFACE/IParameter';


interface IRouteParams {
    id: string;
}

export const DeviationEditForm = (props: ISonaDeviationProps) => {
    {
        const { id } = useParams<IRouteParams>();
        const itemId = id ? Number(id) : 0;
        const history = useHistory();
        const [Stage, setStage] = React.useState(0);
        const [isSubmitting, setIsSubmitting] = React.useState(false);
        const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
        const [Details, setDetails] = React.useState<any[]>([]);
        const [ApprovalMatrixdata, setApprovalMatrix] = React.useState<any[]>([]);
        const [WorkflowHistorydata, setWorkflowHistory] = React.useState<any[]>([]);
        const currentUserRef = React.useRef<any>(null);


        const [ParameterTypeItems, setParameterTypeItems] = React.useState<IParameterType[]>([]);
        const [ParameterTypeOptions, setParameterTypeOptions] = React.useState<IDropdownOption[]>([]);





        //----------------------------     Requestor Details  ---------------------------------


        const [PartNo, setPartNo] = React.useState("");
        const [DeviationNo, setDeviationNo] = React.useState("");

        // ----------------------------- ConcernedDepartment part ---------------------------

        const [ConcernedDepartmentOptions, setConcernedDepartmentOptions] = React.useState<IDropdownOption[]>([]);
        const [selectedConcernedDepartment, setSelectedConcernedDepartment] = React.useState<string | number>();

        // ---------------------------------------------------------------------------

        const [BatchNo, setBatchNo] = React.useState("");

        // ----------------------------- PlantName part ---------------------------

        const [PlantNameOptions, setPlantNameOptions] = React.useState<IDropdownOption[]>([]);
        const [selectedPlantId, setSelectedPlantId] = React.useState<string | number>();
        // const [selectedPlantName, setSelectedPlantName] = React.useState<string>("");

        // ---------------------------------------------------------------------------

        const [ProductionQTY, setProductionQTY] = React.useState("");
        const [RequestorName, setRequestorName] = React.useState("");
        const [ProductionDate, setProductionDate] = React.useState("");

        // ----------------------------- DeviationType part ---------------------------

        const [SelectedDeviationType, setSelectedDeviationType] = React.useState<number | undefined>();

        // ---------------------------------------------------------------------------
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
        const [Status, setStatus] = React.useState("");

        // ------------------------------------------------------------------------------------


        const [DeviationAttachments, setDeviationAttachments] = React.useState<{
            deviationFile?: any;
        }>({});

        const [DeviationAttachFile, setDeviationAttachFile] = useState<File | null>(null);


        const [QualityApprover, setQualityApprover] = React.useState(false);
        const [EngineeringApprover, setEngineeringApprover] = React.useState(false);
        const [COOApprover, setCOOApprover] = React.useState(false);
        const [PlantHeadApprover, setPlantHeadApprover] = React.useState(false);
        const [NextsaApprover, setNextsaApprover] = React.useState(false);
        const [SpecialApproval, setSpecialApproval] = React.useState("");
        const [Comments, setComments] = React.useState("");

        const approverJson = React.useRef<any[]>(null);

        const [WorkflowJSX, setWorkflowJSX] = React.useState(null);


        // --------------------------   table part    ----------------------------

        const [rows, setRows] = useState<any[]>([]);
        const [ParameterOptions, setParameterOptions] = React.useState<IDropdownOption[]>([]);

        // -----------------------------------------------------------------------

        React.useEffect(() => {
            ConcernedDepartmentChoices();
            PlantNameChoices();
            ParameterChoices();
            sp.setup({
                spfxContext: props.currentSPContext
            });

            // ------------------------------------------- Parameter part ---------------------------------------

            ParameterTypeOps().getTopParameterType("Id,ParameterType,SubParameter", "",
                "", { column: 'Created', isAscending: false }, 5000, props)
                .then(results => {
                    setParameterTypeItems(results);
                    const uniqueParameterTypes = Array.from(
                        new Map(results.map(item => [item.ParameterType, item])).values()
                    );
                    const options: IDropdownOption[] = uniqueParameterTypes.map(item => ({
                        key: item.ParameterType,   // ✅ FIX
                        text: item.ParameterType
                    }));
                    setParameterTypeOptions(options);
                });





        }, []);

        const displayWorkflow = () => {

            if (!approverJson.current || approverJson.current.length === 0) {
                setWorkflowJSX(null);
                return;
            }

            const wf: JSX.Element[] = [];

            let isActive;
            let notActive = false;

            approverJson.current.filter(m => m.required === true).forEach((m, i) => {

                if (notActive === false && Stage !== 99) {
                    if (Stage === i) {
                        isActive = 'activeApprover';
                        notActive = true;
                    } else {
                        isActive = 'beforeactiveApprover';
                    }
                } else {
                    isActive = 'overrideStage';
                }

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

        React.useEffect(() => {
            if (!props.currentSPContext) return;

            sp.setup({
                spfxContext: props.currentSPContext
            });
            const loadCurrentUser = async () => {
                try {
                    const user = await sp.web.currentUser();
                    currentUserRef.current = user;
                } catch (error) {
                    console.error("Error fetching current user:", error);
                }
            };
            loadCurrentUser();
        }, [props.currentSPContext]);

        const NPDProductionOptions: IDropdownOption[] = [
            { key: "NPD", text: "NPD" },
            { key: "Production", text: "Production" },
        ];

        const handleChange = (index: number, field: string, value: any) => {
            const updatedRows = [...rows];
            updatedRows[index][field] = value;
            setRows(updatedRows);
        };

        // ✅ Add Row
        const addRow = () => {
            const newRow = {
                id: rows.length + 1,
                parameters: "",
                subParameters: "",
                subParamOptions: [],
                specification: "",
                observation: "",
                date: "",
                quantity: "",
                rootCause: "",
                correctiveAction: "",
                others: "",
                showOthers: false
            };
            setRows([...rows, newRow]);
        };

        const deleteRow = (index: number) => {
            const updatedRows = [...rows];
            updatedRows.splice(index, 1);

            // Optional: prevent deleting last row
            if (updatedRows.length === 0) {
                updatedRows.push({
                    id: 1,
                    parameters: "",
                    specification: "",
                    observation: "",
                    date: "",
                    quantity: "",
                    rootCause: "",
                    correctiveAction: ""
                });
            }

            setRows(updatedRows);
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

        React.useEffect(() => {

            if (
                !Details ||
                Details.length === 0 ||
                !ParameterTypeItems ||
                ParameterTypeItems.length === 0
            ) return;

            const mappedRows = Details.map(
                (d: any, index: number) => {

                    const filteredSubTypes: IDropdownOption[] =
                        ParameterTypeItems
                            .filter(
                                item =>
                                    item.ParameterType === d.Parameters &&
                                    item.SubParameter
                            )
                            .map(item => ({
                                key: item.SubParameter,
                                text: item.SubParameter
                            }));

                    return {
                        id: index + 1,
                        parameters: d.Parameters || "",
                        subParameters: d.SubParameters || "",
                        subParamOptions: filteredSubTypes,
                        specification: d.Specification || "",
                        observation: d.Observation || "",
                        date: d.Date
                            ? d.Date.split("T")[0]
                            : "",
                        quantity: d.Quantity || "",
                        rootCause: d.RootCause || "",
                        correctiveAction: d.CorrectiveAction || "",
                        others: d.others || "",
                        showOthers: false
                    };
                }
            );

            setRows(mappedRows);

        }, [Details, ParameterTypeItems]);

        const loadItem = async (itemId: number) => {
            try {
                const item = await sp.web.lists
                    .getByTitle("DeviationDetails").items.getById(itemId)
                    .select("ID", "PartNo", "DeviationNo", "SupplierName", "BatchNo",
                        "PlantName", "ProductionQTY", "RequestorName", "ProductionDate",
                        "Details", "ApprovalMatrix", "Stage", "WorkflowHistory", "DeviationType",
                        "Remark", "Functional", "Fitment", "SpecialNote", "DeviationRecommended", "QualityName", "QualityProductionDate", "QualityDate",
                        "EngineeringRemark", "EngineeringFunctional", "EngineeringFitment", "EngineeringSpecialNote", "EngineeringDeviationRecommended",
                        "EngineeringName", "EngineeringProductionDate", "EngineeringDate", "SpecialApproval",
                        "PlantHeadRemark", "PlantHeadDeviationAccepted", "PlantHeadProductionDate", "COORemark",
                        "PlantHeadDate", "PlantHeadEffectiveness", "NextApprover/ID", "AttachmentFiles", "Comments"
                    ).expand("NextApprover,AttachmentFiles")();
                setPartNo(item.PartNo || "");
                setDeviationNo(item.DeviationNo || "");
                setSelectedConcernedDepartment(item.SupplierName || "");
                setBatchNo(item.BatchNo || "");
                setSelectedPlantId(item.PlantName || "");
                setSelectedDeviationType(item.DeviationType || "");
                setProductionQTY(item.ProductionQTY || "");
                setRequestorName(item.RequestorName || "");
                setProductionDate(item.ProductionDate ? item.ProductionDate.split("T")[0] : "");
                setStage(item.Stage || "");
                setComments(item.Comments || "");
                setDeviationAttachments({
                    deviationFile: item.AttachmentFiles?.[0] || null
                });

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
                setSpecialApproval(item.SpecialApproval || "");
                approverJson.current = JSON.parse(item.ApprovalMatrix);
                displayWorkflow();

                let parsedDetails: any[] = [];

                if (item.Details) {
                    try {
                        parsedDetails = JSON.parse(item.Details);
                    } catch (parseError) {
                        console.error("Error parsing Details JSON:", parseError);
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

        const deleteAttachment = async (file: any, type: string) => {

            if (!window.confirm("Are you sure you want to delete this attachment?")) return;

            try {
                const itemIdNum = Number(id); // make sure you have item ID

                await sp.web.lists
                    .getByTitle("DeviationDetails")
                    .items.getById(itemIdNum)
                    .attachmentFiles
                    .getByName(file.FileName) // 👈 IMPORTANT
                    .delete();

                // Update UI
                setDeviationAttachments((prev: any) => ({
                    ...prev,
                    [type]: null
                }));

                setDeviationAttachFile(null); // optional cleanup

            } catch (error) {
                console.error("Delete failed", error);
            }
        };

        const loadWorkflow = async (plant?: number, deviationType?: number) => {

            if (!plant || !deviationType) {
                setWorkflowJSX(null);
                return;
            }

            await fetchApprovalMatrix(plant, deviationType);
            displayWorkflow();
        };

        const fetchApprovalMatrix = async (plant?: number, deviationType?: number) => {

            const spCrudOps = await SPCRUDOPS();

            let filterQuery = `Status eq 'Active'`;

            const parentItems = await spCrudOps.getData(
                "DeviationApprovalMatrix",
                "Id,Role/RoleName,Level/Level,Level/Stage,Approver/Title,PlantName/ID,PlantName/PlantName,DeviationType,Approver/ID,Approver/EMail",
                "Role,Level,Approver,PlantName",
                filterQuery,
                { column: "ID", isAscending: true },
                5000,
                props
            );

            let matrix: any[] = [];
            matrix.push(
                {
                    Role: 'Initiator',
                    User: currentUserRef.current?.Title || "",
                    UserID: currentUserRef.current?.Id || "",
                    UserEmail: currentUserRef.current?.Email || "",
                    PendingText: 'Level 0',
                    Stage: 0,
                    required: true
                }
            )
            let COO = parentItems
            .filter(m => (m.Role.RoleName === 'COO'))
            .sort((a, b) => a.Level.Stage - b.Level.Stage); // ✅ SORT BY STAGE
            matrix.push(
                {
                    Role: 'COO',
                    User: COO[0].Approver.Title,
                    UserID: COO[0].Approver.ID,
                    UserEmail: COO[0].Approver.EMail,
                    PendingText: COO[0].Level.Level,
                    Stage: COO[0].Level.Stage,
                    required: false
                }
            )
            let filtermatrix = parentItems
            .filter(m => (m?.PlantName?.PlantName === plant && m?.DeviationType === deviationType))
            .sort((a, b) => a.Level.Stage - b.Level.Stage); // ✅ SORT BY STAGE

            filtermatrix.forEach((item: any, index: number) => {

                matrix.push({
                    Role: item.Role.RoleName,
                    User: item.Approver.Title,
                    UserID: item.Approver.ID,
                    UserEmail: item.Approver.EMail,
                    PendingText: item.Level.Level,
                    Stage: item.Level.Stage,
                    required: true
                });
            });
            approverJson.current = matrix;
            //approvalMatrix.current = parentItems;        
        };

        const ConcernedDepartmentChoices = async () => {
            try {
                const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);

                const field: any = await web.lists
                    .getByTitle("DeviationDetails")
                    .fields
                    .getByInternalNameOrTitle("SupplierName")();

                const choices = field.Choices || [];

                const options = choices.map((choice: string) => ({
                    key: choice,
                    text: choice
                }));

                setConcernedDepartmentOptions(options);

            } catch (error) {
                console.error("Error loading ConcernedDepartment choices:", error);
            }
        };

        const PlantNameChoices = async () => {
            try {
                const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);

                const items: any[] = await web.lists.getByTitle("PlantMaster")
                    .items.select("Id", "PlantName").top(5000)();
                const options: IDropdownOption[] = items.map((item: any) => ({
                    key: item.PlantName,              // ✅ MUST BE NUMBER (ID)
                    text: item.PlantName
                }));

                setPlantNameOptions(options);

            } catch (error) {
                console.error("Error loading PlantName data:", error);
            }
        };

        const ParameterChoices = async () => {
            try {
                const spCrud = await SPCRUDOPS();
                const data = await spCrud.getData("Parameter", "Id,ParameterType", "", "", { column: "Title", isAscending: true }, 5000, props);
                const options = data.map((item: any) => ({
                    key: item.ParameterType,
                    text: item.ParameterType
                }));

                setParameterOptions(options);

            } catch (error) {
                console.error("Error loading Parameter:", error);
            }
        };

        const handleUpdate = async () => {

            // 🔹 Field validations

            if (!PartNo) {
                alert("Please enter the Part Number");
                return;
            }

            if (!DeviationNo) {
                alert("Please enter the Deviation Number");
                return;
            }

            if (!selectedConcernedDepartment) {
                alert("Please select the Concerned Department");
                return;
            }

            if (!BatchNo) {
                alert("Please enter the Batch/Lot Number");
                return;
            }
            if (!selectedPlantId) {
                alert("Please select the Plant Name");
                return;
            }

            if (!ProductionQTY) {
                alert("Please enter the Production Quantity");
                return;
            }

            if (!ProductionDate) {
                alert("Please enter your Production Date");
                return;
            }

            if (!SelectedDeviationType) {
                alert("Please select the Deviation Type");
                return;
            }

            if (!DeviationAttachFile && !DeviationAttachments?.deviationFile) {
                alert("Please upload Attachment");
                return;
            }

            // 🚫 prevent double submission
            if (isSubmitting) return;
            setIsSubmitting(true);

            const itemId = Number(id);

            if (!itemId) {
                alert("Invalid item ID");
                setIsSubmitting(false);
                return;
            }

            const now = new Date();
            const formattedDate = now.toLocaleDateString("en-GB");

            const currentUser = await sp.web.currentUser();

            try {


                const productionQtyNum = Number(ProductionQTY || 0);
                let errorMessages: string[] = [];
                for (let i = 0; i < rows.length; i++) {
                    const qty = Number(rows[i].quantity || 0);

                    if (qty > productionQtyNum) {
                        errorMessages.push(
                            `Row ${i + 1}: Quantity (${qty}) cannot exceed Production QTY (${productionQtyNum})`
                        );
                    }
                }
                if (errorMessages.length > 0) {
                    alert(errorMessages.join("\n"));
                    setIsSubmitting(false);
                    return;
                }

                // ✅ Prepare table data (rows)
                const tableData = rows.map((row, index) => ({
                    SNo: index + 1,
                    Parameters: row.parameters,
                    SubParameters: row.subParameters, // ✅ ADD THIS
                    others: row.others,
                    Specification: row.specification,
                    Observation: row.observation,
                    Date: row.date,
                    Quantity: row.quantity,
                    RootCause: row.rootCause,
                    CorrectiveAction: row.correctiveAction
                }));

                let tabledata = JSON.stringify(tableData);
                let apprid = approverJson.current.filter(m => m.required === true).filter(m => m.Stage === 1)[0]?.UserID || "";

                let workflowHistory = [...WorkflowHistorydata];
                const newHistoryEntry = {
                    CurrentApprover: currentUser.Title,
                    ActionTaken: 'Request submitted',
                    Date: formattedDate,
                    CurrentStatus: 'Submitted to Engineering'
                };
                workflowHistory.push(newHistoryEntry);
                const updatedHistory = JSON.stringify(workflowHistory);

                await sp.web.lists
                    .getByTitle("DeviationDetails")
                    .items.getById(itemId)
                    .update({
                        Details: tabledata,
                        PartNo: PartNo,
                        DeviationNo: DeviationNo,
                        SupplierName: selectedConcernedDepartment,
                        BatchNo: BatchNo,
                        PlantName: selectedPlantId,
                        ProductionQTY: ProductionQTY,
                        RequestorName: currentUserRef.current?.Title || "",
                        ProductionDate: ProductionDate ? new Date(ProductionDate) : null,
                        DeviationType: SelectedDeviationType,
                        ApprovalMatrix: JSON.stringify(approverJson.current),
                        WorkflowHistory: updatedHistory,
                        Stage: 1,
                        Status: "Pending for Approval",
                        NextApproverId: apprid || ""
                    });

                if (DeviationAttachFile) {
                    await sp.web.lists
                        .getByTitle("DeviationDetails")
                        .items.getById(itemId)
                        .attachmentFiles.add(
                            DeviationAttachFile.name,
                            DeviationAttachFile
                        );
                }

                alert("Request submitted successfully!");

                history.push("/");


            } catch (error) {

                console.error("Error updating item:", error);

                alert("Update failed. Please try again.");

                setIsSubmitting(false);

            }
        };

        const showOthersColumn = Details.some(r => r.others && r.others.trim() !== "");

        return (
            <div>
                <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                    <div className='row'>
                        <div className='col-md-12'>
                            <div className='Main-Boxpoup'>
                                <div className="bordered">
                                    <a><img src={logo} /></a>
                                    <h1>Deviation Edit Form</h1>
                                </div>
                                <div className='displayWF'>{WorkflowJSX}</div>
                                <div className='borderedbox'>

                                    <div className="heading1" style={{ marginTop: "10px" }}>
                                        <label>Requestor Information</label>
                                    </div>

                                    <div className='main-formcontainer'>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Part No <span className='Mantorystar'>*</span></label>
                                                <input type="text" className='form-control' value={PartNo} onChange={e => setPartNo(e.target.value)} />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>Deviation No <span className='Mantorystar'>*</span></label>
                                                <input type="text" className='form-control' value={DeviationNo} onChange={e => setDeviationNo(e.target.value)} />
                                            </div>

                                            <div className='col-md-4'>
                                                <label className='font'>Concerned Department <span className='Mantorystar'>*</span></label>
                                                <Dropdown
                                                    options={ConcernedDepartmentOptions}
                                                    selectedKey={selectedConcernedDepartment}
                                                    onChange={(e, option) => setSelectedConcernedDepartment(option?.key as number)}
                                                />
                                            </div>

                                        </div>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Batch / Lot No <span className='Mantorystar'>*</span></label>
                                                <input type="text" className='form-control' value={BatchNo} onChange={e => setBatchNo(e.target.value)} />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>Plant Name <span className='Mantorystar'>*</span></label>
                                                <Dropdown
                                                    options={PlantNameOptions}
                                                    selectedKey={selectedPlantId}
                                                    onChange={async (e, option) => {
                                                        setSelectedPlantId(option?.key as number);
                                                        // setSelectedPlantName(option?.text as string);
                                                        await loadWorkflow(option?.key as number, SelectedDeviationType);
                                                    }}
                                                />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>Production QTY <span className='Mantorystar'>*</span></label>
                                                <input type='text' className='form-control' value={ProductionQTY} onChange={(e) => { const value = e.target.value.replace(/[^0-9]/g, ""); setProductionQTY(value); }} />
                                            </div>
                                        </div>
                                        <div className='row mb-20'>
                                            <div className='col-md-4'>
                                                <label className='font'>Requester Name <span className='Mantorystar'>*</span></label>
                                                <input type="text" value={RequestorName} readOnly className='form-control readonly' />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>Production Date <span className='Mantorystar'>*</span></label>
                                                <input type="date" className='form-control' value={ProductionDate} onChange={e => setProductionDate(e.target.value)} />
                                            </div>
                                            <div className='col-md-4'>
                                                <label className='font'>Deviation Type <span className='Mantorystar'>*</span></label>
                                                <Dropdown
                                                    options={NPDProductionOptions}
                                                    selectedKey={SelectedDeviationType}
                                                    onChange={async (e, option) => {
                                                        const value = option?.key as number;
                                                        setSelectedDeviationType(value);
                                                        // await loadWorkflow(selectedPlantId, value);
                                                    }}
                                                />
                                            </div>
                                        </div>
                                    </div>

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
                                                    <col style={{ width: "90px" }} />
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
                                                        <th>Add</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {rows.length === 0 ? (
                                                        <tr>
                                                            <td colSpan={9} style={{ textAlign: "center" }}>
                                                                Loading...
                                                            </td>
                                                        </tr>
                                                    ) : (
                                                        rows.map((row, index) => (
                                                            <tr key={row.id}>
                                                                <td>{index + 1}</td>
                                                                <td>
                                                                    {/* <Dropdown
                                                                        options={ParameterOptions}
                                                                        selectedKey={row.parameters}
                                                                        onChange={(e, option) =>
                                                                            handleChange(index, "parameters", option?.key)
                                                                        }
                                                                        style={{ fontSize: "13px" }}
                                                                    /> */}
                                                                    <Dropdown
                                                                        options={ParameterTypeOptions}
                                                                        selectedKey={row.parameters}
                                                                        onChange={(e, option) => {

                                                                            const selectedText =
                                                                                option?.text as string;

                                                                            const filteredSubTypes =
                                                                                ParameterTypeItems
                                                                                    .filter(
                                                                                        item =>
                                                                                            item.ParameterType === selectedText &&
                                                                                            item.SubParameter
                                                                                    )
                                                                                    .map(item => ({
                                                                                        key: item.SubParameter,
                                                                                        text: item.SubParameter
                                                                                    }));

                                                                            handleChange(
                                                                                index,
                                                                                "parameters",
                                                                                selectedText
                                                                            );

                                                                            handleChange(
                                                                                index,
                                                                                "subParameters",
                                                                                undefined
                                                                            );

                                                                            handleChange(
                                                                                index,
                                                                                "subParamOptions",
                                                                                filteredSubTypes
                                                                            );
                                                                        }}
                                                                    />
                                                                </td>
                                                                <td>
                                                                    <Dropdown
                                                                        options={row.subParamOptions || []}
                                                                        selectedKey={row.subParameters}
                                                                        onChange={(e, option) =>
                                                                            handleChange(
                                                                                index,
                                                                                "subParameters",
                                                                                option?.key
                                                                            )
                                                                        }
                                                                    />
                                                                </td>
                                                                {showOthersColumn && (
                                                                    <td>
                                                                        <textarea
                                                                            className='form-control'
                                                                            style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }}
                                                                            value={row.others}
                                                                            onChange={(e) => 
                                                                                handleChange(index, 'others', e.currentTarget.value)
                                                                            }
                                                                        />
                                                                    </td>
                                                                )}

                                                                <td>
                                                                    <textarea
                                                                        className='form-control'
                                                                        style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }}
                                                                        value={row.specification}
                                                                        onChange={(e) => {
                                                                            const textarea = e.currentTarget;
                                                                            textarea.style.height = "auto";
                                                                            textarea.style.height = textarea.scrollHeight + "px";
                                                                            handleChange(index, "specification", textarea.value);
                                                                        }}
                                                                    />
                                                                </td>

                                                                <td>
                                                                    <textarea
                                                                        className='form-control'
                                                                        style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }}
                                                                        value={row.observation}
                                                                        onChange={(e) => {
                                                                            const textarea = e.currentTarget;
                                                                            textarea.style.height = "auto";
                                                                            textarea.style.height = textarea.scrollHeight + "px";
                                                                            handleChange(index, "observation", textarea.value);
                                                                        }}
                                                                    />
                                                                </td>

                                                                <td>
                                                                    <TextField
                                                                        type="date"
                                                                        value={row.date}
                                                                        onChange={(e) =>
                                                                            handleChange(index, "date", e.currentTarget.value)
                                                                        }
                                                                        style={{ fontSize: "13px" }}
                                                                    />
                                                                </td>

                                                                <td>
                                                                    <TextField
                                                                        value={row.quantity}
                                                                        onKeyPress={(e) => {
                                                                            if (!/[0-9]/.test(e.key)) e.preventDefault();
                                                                        }}
                                                                        onChange={(e) =>
                                                                            handleChange(index, "quantity", e.currentTarget.value)
                                                                        }
                                                                        style={{ fontSize: "13px" }}
                                                                    />
                                                                </td>

                                                                <td>
                                                                    <textarea
                                                                        className='form-control'
                                                                        style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }}
                                                                        value={row.rootCause}
                                                                        onChange={(e) => {
                                                                            const textarea = e.currentTarget;
                                                                            textarea.style.height = "auto";
                                                                            textarea.style.height = textarea.scrollHeight + "px";
                                                                            handleChange(index, "rootCause", textarea.value);
                                                                        }}
                                                                    />
                                                                </td>

                                                                <td>
                                                                    <textarea
                                                                        className='form-control'
                                                                        style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }}
                                                                        value={row.correctiveAction}
                                                                        onChange={(e) => {
                                                                            const textarea = e.currentTarget;
                                                                            textarea.style.height = "auto";
                                                                            textarea.style.height = textarea.scrollHeight + "px";
                                                                            handleChange(index, "correctiveAction", textarea.value);
                                                                        }}
                                                                    />
                                                                </td>

                                                                <td>
                                                                    <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: "5px" }}>
                                                                        <DefaultButton onClick={addRow} iconProps={{ iconName: 'Add' }} className="Add" />
                                                                        <DefaultButton onClick={() => deleteRow(index)} iconProps={{ iconName: 'Delete' }} className="Delete" />
                                                                    </div>
                                                                </td>
                                                            </tr>
                                                        ))
                                                    )}
                                                </tbody>
                                            </table>
                                        </div>
                                    </div>

                                    {((COOApprover || QualityApprover || PlantHeadApprover) || (Status === 'Rejected' || Status === 'Send Back')) && (
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
                                            </div>
                                        </>
                                    )}

                                    {(((QualityApprover || PlantHeadApprover) && SpecialApproval === "Yes") || ((Status === 'Rejected' || Status === 'Send Back') && COORemark)) && (
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

                                    {((PlantHeadApprover && SpecialApproval === "No") || ((Status === 'Rejected' || Status === 'Send Back') && DeviationRecommended && Functional && Remark && Fitment && SpecialNote)) && (
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
                                        <label>Upload Documents</label>
                                    </div>
                                    <div className="main-formcontainer">
                                        <div className='row mb-20'>
                                            <div className='col-md-3'>
                                                <label className='font'>Attachment <span className='Mantorystar'>*</span></label>
                                                {DeviationAttachments.deviationFile ? (
                                                    <div className="d-flex align-items-center gap-2 mt-1">
                                                        {/* File Name */}
                                                        <a href={DeviationAttachments.deviationFile.ServerRelativeUrl} target="_blank" rel="noopener noreferrer">
                                                            {DeviationAttachments.deviationFile.FileName}
                                                        </a>
                                                        {/* Cross Icon */}
                                                        <span style={{ color: "red", fontWeight: "bold", cursor: "pointer", fontSize: "18px" }}
                                                            title="Delete attachment"
                                                            onClick={() => deleteAttachment(DeviationAttachments.deviationFile, "deviationFile")}>
                                                            ×
                                                        </span>
                                                    </div>
                                                ) : (
                                                    <Upload
                                                        beforeUpload={(file) => { setDeviationAttachFile(file); return false; }}
                                                        className="upload-full-width">
                                                        <Button className="upload-btn-full" icon={<UploadOutlined />} iconPosition="end"></Button>
                                                    </Upload>
                                                )}
                                            </div>
                                        </div>
                                    </div>
                                    <div className='row my-3'>
                                        <div className='col-md-12'>
                                            <div style={{ display: "flex", justifyContent: "center", gap: "5px" }}>
                                                <a onClick={() => !isSubmitting && handleUpdate()} className="submit-btn">
                                                    Submit
                                                </a>
                                                {/* <a onClick={() => !isSubmitting && handleSaveDraft(itemId)}  className="Rework-btn">
                                                    Save as Draft
                                                </a> */}
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

            </div>
        )
    }
}