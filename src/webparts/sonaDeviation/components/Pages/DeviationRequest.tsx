import * as React from 'react';
import { useState, useEffect } from 'react';
import type { ISonaDeviationProps } from '../ISonaDeviationProps';
import { TextField, DefaultButton, Dropdown, IDropdownOption } from '@fluentui/react';
import { Link, useHistory } from "react-router-dom";
import logo from "../../assets/sona-comstarlogo.png";
// import 'bootstrap/dist/css/bootstrap.min.css';
import "../Pages/Css/NewRequest.scss";

import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";


import { Web } from "@pnp/sp/webs";

import { Upload, Button } from "antd";
import { UploadOutlined } from "@ant-design/icons";

import SPCRUDOPS from "../../service/BAL/spcrud";

import ParameterTypeOps from '../../service/BAL/Parameter';
import { IParameterType } from '../../service/INTERFACE/IParameter';

import UserProfileOps from '../../service/BAL/UserProfile';
import { IUserProfile } from '../../service/INTERFACE/IUserProfile';

export const NewRequest = (props: ISonaDeviationProps) => {

    const history = useHistory();
    const web = Web(props.currentSPContext.pageContext.web.absoluteUrl);
    const approverJson = React.useRef<any[]>(null);
    const approvalMatrix = React.useRef<any[]>(null);
    const [isSubmitting, setIsSubmitting] = useState(false);

    const [PartNo, setPartNo] = React.useState("");

    // For Counter 
    const [counterCode, setcounterCode] = useState<any[]>([]);

    // const [DeviationNo, setDeviationNo] = React.useState("");

    const [ConcernedDepartmentOptions, setConcernedDepartmentOptions] = React.useState<IDropdownOption[]>([]);
    const [selectedConcernedDepartment, setSelectedConcernedDepartment] = React.useState<string | number>();

    const [BatchNo, setBatchNo] = React.useState("");

    const [PlantNameOptions, setPlantNameOptions] = React.useState<IDropdownOption[]>([]);

    const [ProductionQTY, setProductionQTY] = React.useState("");

    const currentUserRef = React.useRef<any>(null);
    const [ProductionDate, setProductionDate] = React.useState("");


    const [SelectedDeviationType, setSelectedDeviationType] = React.useState<number | undefined>();

    const [selectedPlantId, setSelectedPlantId] = React.useState<number | undefined>();
    const [selectedPlantName, setSelectedPlantName] = React.useState<string>("");

    const [userprofile, setuserprofile] = React.useState<IUserProfile | null>(null);
    const [departmentmasterdata, setDepartmentmasterdata] = React.useState<any[]>([]);
    const [Locationdata, setLocationdata] = React.useState("");

    const [WorkflowJSX, setWorkflowJSX] = React.useState(null);
    const [Stage, setStage] = React.useState(0);

    const [DeviationAttachFile, setDeviationAttachFile] = useState<File | null>(null);
    // const [ParameterypeTOptions, setParameterTypeOptions] = React.useState<IDropdownOption[]>([]);
    // const [selectedParameter, setSelectedParameter] = React.useState<string | number>();




    const [ParameterTypeItems, setParameterTypeItems] = React.useState<IParameterType[]>([]);
    const [ParameterTypeOptions, setParameterTypeOptions] = React.useState<IDropdownOption[]>([]);
    const [selectedParameterType, setSelectedParameterType] = React.useState<string | number>();

    const [SubParameterTypeOptions, setSubParameterTypeOptions] = React.useState<IDropdownOption[]>([]);
    const [selectedSubParameterType, setSelectedSubParameterType] = React.useState<string | number>();



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
        ConcernedDepartmentChoices();
        PlantNameChoices();
        // ParameterChoices();
        userProfile();
        sp.setup({
            spfxContext: props.currentSPContext
        });

        // ------------------------------------------- Parameter part ---------------------------------------

        ParameterTypeOps().getTopParameterType("Id,PlantName/Id,PlantName/PlantName,ParameterType,SubParameter", "PlantName",
            "", { column: 'Created', isAscending: false }, 5000, props)
            .then(results => {
                setParameterTypeItems(results);
                // const uniqueParameterTypes = Array.from(
                //     new Map(results.map(item => [item.ParameterType, item])).values()
                // );
                // const options: IDropdownOption[] = uniqueParameterTypes.map(item => ({
                //     key: item.Id,
                //     text: item.ParameterType
                // }));
                // setParameterTypeOptions(options);
            });

    }, []);

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


    const loadWorkflow = async (plant?: number, deviationType?: number) => {

        if (!plant || !deviationType) {
            setWorkflowJSX(null);
            return;
        }

        await fetchApprovalMatrix(plant, deviationType);
        displayWorkflow();
    };

    const NPDProductionOptions: IDropdownOption[] = [
        { key: "NPD", text: "NPD" },
        { key: "Production", text: "Production" },
    ];

    const userProfile = async () => {
        const userprofiledata = await UserProfileOps().getLoggUserProfile(props);
        if (userprofiledata) {
            setuserprofile(userprofiledata);
        }
        setLocationdata(userprofiledata.Location);
        //userdept.current = userprofiledata.UserProfileProperties.filter(m => m?.Key === 'Department')[0]?.Value;
        //return userprofiledata;
    };

    // ✅ Table State
    const [rows, setRows] = useState([
        {
            id: 1,
            parameters: undefined as number | undefined,        // for dropdown selectedKey
            parameterText: "" as string,                        // for saving

            subParameter: undefined as number | undefined,      // for dropdown selectedKey
            subParameterText: "" as string,                     // for saving
            subParameterOptions: [] as IDropdownOption[],
            specification: "",
            observation: "",
            date: "",
            quantity: "",
            rootCause: "",
            correctiveAction: "",
            others: "",
            showOthers: false
        }
    ]);
    const showOthersColumn = rows.some(r => r.showOthers);

    // ✅ Handle Input Change
    const handleChange = (index: number, field: string, value: any) => {
        const updatedRows = [...rows];
        (updatedRows[index] as any)[field] = value;
        setRows(updatedRows);
    };

    // ✅ Add Row
    const addRow = () => {
        setRows([
            ...rows,
            {
                id: rows.length + 1,
                parameters: undefined,
                parameterText: "",
                subParameter: undefined,
                subParameterText: "",
                subParameterOptions: [],
                specification: "",
                observation: "",
                date: "",
                quantity: "",
                rootCause: "",
                correctiveAction: "",
                others: "",
                showOthers: false
            }
        ]);
    };

    const deleteRow = (index: number) => {
        const updatedRows = [...rows];
        updatedRows.splice(index, 1);

        if (updatedRows.length === 0) {
            updatedRows.push({
                id: 1,
                parameters: undefined,
                parameterText: "",
                subParameter: undefined,
                subParameterText: "",
                subParameterOptions: [],
                specification: "",
                observation: "",
                date: "",
                quantity: "",
                rootCause: "",
                correctiveAction: "",
                others: "",
                showOthers: false
            });
        }

        setRows(updatedRows);
    };

    const fetchApprovalMatrix = async (plant?: number, deviationType?: number) => {

        const spCrudOps = await SPCRUDOPS();

        let filterQuery = `Status eq 'Active'`;

        const parentItems = await spCrudOps.getData(
            "DeviationApprovalMatrix",
            "Id,Role/RoleName,Level/Level,Level/Stage,Approver/Title,PlantName/ID,DeviationType,Approver/ID,Approver/EMail",
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
        .filter(m => (m?.PlantName?.ID === plant && m?.DeviationType === deviationType))
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



    // ******************** Deviation no counter set here ***********************

    async function getDeviationRequestCounter() {
        try {
            const spCrudOps = await SPCRUDOPS();
            const counterCode = await spCrudOps.getData(
                "DeviationRequestCounter",
                "ID,FormatRequest,Counter",
                "",
                "",
                { column: "ID", isAscending: true },
                100,
                props
            );

            console.log("Counter Data:", counterCode);
            setcounterCode(counterCode);
        } catch (err) {
            console.error("Error loading customers:", err);
        }
    }

    useEffect(() => {
        getDeviationRequestCounter();
    }, []);

    // ******************** Deviation no counter set end here ***********************

    const getFinancialYear = (): string => {
        const today = new Date();
        const year = today.getFullYear();
        const month = today.getMonth() + 1;

        return month >= 4
            ? `${year}-${(year + 1).toString().slice(2)}`
            : `${year - 1}-${year.toString().slice(2)}`;
    };



    const generateDeviationNumber = async (partNo: string): Promise<string> => {
        const fy = getFinancialYear();

        try {
            const spCrudOps = await SPCRUDOPS();

            const existingItems = await spCrudOps.getData(
                "DeviationDetails",
                "DeviationNo,PartNo,FinancialYear",
                "",
                "",
                { column: "ID", isAscending: false },
                5000,
                props
            );

            // ✅ HANDLE BOTH CASES
            const allItems = Array.isArray(existingItems?.data)
                ? existingItems.data
                : Array.isArray(existingItems)
                    ? existingItems
                    : [];

            console.log("ALL ITEMS:", allItems);

            const filteredItems = allItems.filter((item: any) => {
                return (
                    item.PartNo?.toString().trim() === partNo.toString().trim() &&
                    item.FinancialYear &&
                    item.FinancialYear.toString().trim() === fy.toString().trim()
                );
            });

            console.log("FILTERED ITEMS:", filteredItems);

            let maxSequence = 0;

            filteredItems.forEach((item: any) => {
                const devNo: string = item.DeviationNo;

                if (devNo) {
                    const parts = devNo.split("-");
                    const seq = parseInt(parts[parts.length - 1]);

                    if (!isNaN(seq) && seq > maxSequence) {
                        maxSequence = seq;
                    }
                }
            });

            const nextNumber = maxSequence + 1;
            const sequence = nextNumber.toString().padStart(3, "0");

            return `${fy}-${partNo}-${sequence}`;

        } catch (error) {
            console.error("Error generating DeviationNo:", error);
            throw error;
        }
    };


    // old using deviation no diffrent 
    // const handleSubmit = async () => {

    //     if (!PartNo) {
    //         alert("Please enter the Part Number");
    //         return;
    //     }
    //     // ✅ NEW VALIDATION
    //     if (PartNo.length < 6 || PartNo.length > 17) {
    //         alert("Part Number must be between 6 and 17 characters");
    //         return;
    //     }
    //     // if (!DeviationNo) {
    //     //     alert("Please enter the Deviation Number");
    //     //     return;
    //     // }

    //     if (!selectedConcernedDepartment) {
    //         alert("Please select the Concerned Department");
    //         return;
    //     }

    //     if (!BatchNo) {
    //         alert("Please enter the Batch/Lot Number");
    //         return;
    //     }
    //     if (!selectedPlantName) {
    //         alert("Please select the Plant Name");
    //         return;
    //     }

    //     if (!ProductionQTY) {
    //         alert("Please enter the Production Quantity");
    //         return;
    //     }

    //     if (!ProductionDate) {
    //         alert("Please enter your Production Date");
    //         return;
    //     }

    //     if (!SelectedDeviationType) {
    //         alert("Please select the Deviation Type");
    //         return;
    //     }

    //     if (!DeviationAttachFile) {
    //         alert("Please upload attachment");
    //         return;
    //     }

    //     // 🚫 prevent double submission
    //     if (isSubmitting) return;
    //     setIsSubmitting(true);

    //     const now = new Date();
    //     const formattedDate = now.toLocaleDateString("en-GB");

    //     try {


    //         const fy = getFinancialYear();
    //         const deviationNumber = await generateDeviationNumber(PartNo);

    //         const productionQtyNum = Number(ProductionQTY || 0);
    //         let errorMessages: string[] = [];
    //         for (let i = 0; i < rows.length; i++) {
    //             const qty = Number(rows[i].quantity || 0);

    //             if (qty > productionQtyNum) {
    //                 errorMessages.push(
    //                     `Row ${i + 1}: Quantity (${qty}) cannot exceed Production QTY (${productionQtyNum})`
    //                 );
    //             }
    //         }
    //         if (errorMessages.length > 0) {
    //             alert(errorMessages.join("\n"));
    //             setIsSubmitting(false);
    //             return;
    //         }

    //         // ✅ Prepare table data (rows)
    //         const tableData = rows.map((row, index) => ({
    //             SNo: index + 1,
    //             Parameters: row.parameterText,
    //             SubParameters: row.subParameterText,
    //             Specification: row.specification,
    //             Observation: row.observation,
    //             Date: row.date,
    //             Quantity: row.quantity,
    //             RootCause: row.rootCause,
    //             CorrectiveAction: row.correctiveAction
    //         }));

    //         let tabledata = JSON.stringify(tableData);
    //         let apprid = approverJson.current.filter(m => m.required === true).filter(m => m.Stage === 1)[0]?.UserID || "";

    //         const spCrudOps = await SPCRUDOPS();

    //         const parentItems = await spCrudOps.insertData("DeviationDetails", {
    //             FinancialYear: fy,
    //             Details: tabledata,
    //             PartNo: PartNo,
    //             DeviationNo: deviationNumber,
    //             SupplierName: selectedConcernedDepartment,
    //             BatchNo: BatchNo,
    //             PlantName: selectedPlantName,
    //             ProductionQTY: ProductionQTY,
    //             RequestorName: currentUserRef.current?.Title || "",
    //             ProductionDate: ProductionDate,
    //             DeviationType: SelectedDeviationType,
    //             ApprovalMatrix: JSON.stringify(approverJson.current),
    //             WorkflowHistory: JSON.stringify([
    //                 {
    //                     CurrentApprover: props.currentSPContext.pageContext.user.displayName
    //                         || props.currentSPContext.pageContext.user.email
    //                         || '',
    //                     ActionTaken: 'Request submitted',
    //                     Comment: '',
    //                     Date: formattedDate,
    //                     CurrentStatus: 'Submitted to Engineering'
    //                 }
    //             ]),
    //             Stage: 1,
    //             Status: "Pending for Approval",
    //             NextApproverId: apprid || ""
    //         }, props);

    //         // 🔹 Get created item ID
    //         const itemId = parentItems?.data?.ID;

    //         // 🔹 Upload attachment (FIXED VERSION)
    //         if (DeviationAttachFile && itemId) {
    //             try {
    //                 await sp.web.lists
    //                     .getByTitle("DeviationDetails")
    //                     .items.getById(itemId)
    //                     .attachmentFiles.add(
    //                         DeviationAttachFile.name,
    //                         DeviationAttachFile
    //                     );
    //             } catch (err) {
    //                 console.error("Error uploading attachment:", err);
    //                 alert("Item created but attachment upload failed");
    //             }
    //         }

    //         alert("Request submitted successfully");

    //         history.push("/");

    //     } catch (error) {
    //         console.error("Error submitting Deviation Request:", error);
    //         alert("Error submitting Deviation Request");
    //         setIsSubmitting(false);
    //     }
    // };

    const handleSubmit = async () => {

        if (!PartNo) {
            alert("Please enter the Part Number");
            return;
        }
        // ✅ NEW VALIDATION
        if (PartNo.length < 6 || PartNo.length > 17) {
            alert("Part Number must be between 6 and 17 characters");
            return;
        }
        // if (!DeviationNo) {
        //     alert("Please enter the Deviation Number");
        //     return;
        // }

        if (!selectedConcernedDepartment) {
            alert("Please select the Concerned Department");
            return;
        }

        if (!BatchNo) {
            alert("Please enter the Batch/Lot Number");
            return;
        }
        if (!selectedPlantName) {
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

        if (!DeviationAttachFile) {
            alert("Please upload attachment");
            return;
        }

        // 🚫 prevent double submission
        if (isSubmitting) return;
        setIsSubmitting(true);

        const now = new Date();
        const formattedDate = now.toLocaleDateString("en-GB");

        try {
            // ✅ Generate Deviation Number
            const fy = getFinancialYear();
            const year = new Date().getFullYear();

            const currentCounter = Number(counterCode[0]?.Counter || 0);
            const nextCounter = currentCounter + 1;

            const paddedCounter = String(nextCounter).padStart(4, "0");

            const deviationNumber = `FY-${year}-${paddedCounter}`;

            const productionQtyNum = Number(ProductionQTY || 0);
            // ✅ Validate table rows
            let tableErrors: string[] = [];
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                if (!row.parameterText) {
                    tableErrors.push(`Row ${i + 1}: Please select Parameter`);
                }
                if (!row.subParameterText) {
                    tableErrors.push(`Row ${i + 1}: Please select Sub Parameter`);
                }
                if (!row.date) {
                    tableErrors.push(`Row ${i + 1}: Please enter Processing Date`);
                }
                if (!row.quantity) {
                    tableErrors.push(`Row ${i + 1}: Please enter Quantity`);
                }
                const qty = Number(row.quantity || 0);
                if (qty > productionQtyNum) {
                    tableErrors.push(
                        `Row ${i + 1}: Quantity (${qty}) cannot exceed Production QTY (${productionQtyNum})`
                    );
                }
            }
            // ✅ Stop submission if validation fails
            if (tableErrors.length > 0) {
                alert(tableErrors.join("\n"));
                setIsSubmitting(false);
                return;
            }

            // ✅ Prepare table data (rows)
            const tableData = rows.map((row, index) => ({
                SNo: index + 1,
                Parameters: row.parameterText,
                SubParameters: row.subParameterText,
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

            const spCrudOps = await SPCRUDOPS();

            const parentItems = await spCrudOps.insertData("DeviationDetails", {
                FinancialYear: fy,
                Details: tabledata,
                PartNo: PartNo,
                DeviationNo: deviationNumber,
                SupplierName: selectedConcernedDepartment,
                BatchNo: BatchNo,
                PlantName: selectedPlantName,
                // ProductionQTY: ProductionQTY,
                ProductionQTY: ProductionQTY,
                RequestorName: currentUserRef.current?.Title || "",
                ProductionDate: ProductionDate,
                DeviationType: SelectedDeviationType,
                ApprovalMatrix: JSON.stringify(approverJson.current),
                WorkflowHistory: JSON.stringify([
                    {
                        CurrentApprover: props.currentSPContext.pageContext.user.displayName
                            || props.currentSPContext.pageContext.user.email
                            || '',
                        ActionTaken: 'Request submitted',
                        Comment: '',
                        Date: formattedDate,
                        CurrentStatus: 'Submitted to Engineering'
                    }
                ]),
                Stage: 1,
                Status: "Pending for Approval",
                NextApproverId: apprid || null
                // NextApproverId: apprid ? Number(apprid) : null
            }, props);

            // 🔹 Get created item ID
            const itemId = parentItems?.data?.ID;

            // 🔹 Upload attachment (FIXED VERSION)
            if (DeviationAttachFile && itemId) {
                try {
                    await sp.web.lists
                        .getByTitle("DeviationDetails")
                        .items.getById(itemId)
                        .attachmentFiles.add(
                            DeviationAttachFile.name,
                            DeviationAttachFile
                        );
                } catch (err) {
                    console.error("Error uploading attachment:", err);
                    alert("Item created but attachment upload failed");
                }
            }

            // For Counter Update **************
            await spCrudOps.updateData(
                "DeviationRequestCounter",
                counterCode[0].ID,
                {
                    Counter: String(nextCounter)
                },
                props
            );

            // For Counter Update **************

            alert("Request submitted successfully");

            history.push("/");

        } catch (error) {
            console.error("Error submitting Deviation Request:", error);
            alert("Error submitting Deviation Request");
            setIsSubmitting(false);
        }
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

            const items: any[] = await web.lists
                .getByTitle("PlantMaster")
                .items
                .select("Id", "PlantName")
                .top(5000)();

            const options: IDropdownOption[] = items.map((item: any) => ({
                key: item.Id,              // ✅ MUST BE NUMBER (ID)
                text: item.PlantName
            }));

            setPlantNameOptions(options);

        } catch (error) {
            console.error("Error loading PlantName data:", error);
        }
    };

    // const ParameterChoices = async () => {
    //     try {
    //         const spCrud = await SPCRUDOPS();
    //         const data = await spCrud.getData("Parameter",
    //             "Id,ParameterType", "", "", { column: "Title", isAscending: true }, 5000, props);
    //         const options = data.map((item: any) => ({
    //             key: item.ParameterType,
    //             text: item.ParameterType
    //         }));
    //         setParameterOptions(options);
    //     } catch (error) {
    //         console.error("Error loading Parameter:", error);
    //     }
    // };


    return (
        <div>

            {/* 🔹 Loader */}
            {isSubmitting && (
                <div className="loading-overlay">
                    <div className="loading-content">
                        <p>Please wait, submitting request...</p>
                    </div>
                </div>
            )}

            <div className='MainUplodForm' style={{ margin: "5px 0px" }}>
                <div className='row'>
                    <div className='col-md-12'>
                        <div className='Main-Boxpoup'>

                            {/* 🔹 Header */}
                            <div className="bordered">
                                <img src={logo} />
                                <h1>REQUEST FOR DEVIATION </h1>
                            </div>
                            <div className='displayWF'>{WorkflowJSX}</div>
                            <div className='borderedbox'>
                                {/* 🔹 Section Title */}
                                <div className="heading1">
                                    <label>Requestor Information</label>
                                </div>
                                <div className='main-formcontainer'>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label htmlFor="Part No" className='font'>Part No <span className="Mantorystar">*</span></label>
                                            <input type="text" className='form-control' onChange={e => setPartNo(e.target.value)} />
                                        </div>
                                        {/* <div className='col-md-4'>
                                            <label htmlFor="Deviation No" className='font'>Deviation No <span className="Mantorystar">*</span></label>
                                            <input type="text" className='form-control readonly' onChange={e => setDeviationNo(e.target.value)} />
                                        </div> */}
                                        <div className='col-md-4'>
                                            <label htmlFor="Concerned Department" className='font'>Concerned Department <span className="Mantorystar">*</span></label>
                                            <Dropdown
                                                options={ConcernedDepartmentOptions}
                                                selectedKey={selectedConcernedDepartment}
                                                onChange={(e, option) => setSelectedConcernedDepartment(option?.key)}
                                            />
                                        </div>
                                    </div>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label htmlFor="Batch / Lot No" className='font'>Batch / Lot No <span className="Mantorystar">*</span></label>
                                            <input type="text" className='form-control' onChange={e => setBatchNo(e.target.value)} />
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="Plant Name" className='font'>Plant Name <span className="Mantorystar">*</span></label>
                                            <Dropdown
                                                options={PlantNameOptions}
                                                selectedKey={selectedPlantId} // ✅ MUST MATCH key
                                                onChange={async (e, option) => {
                                                    const plantId = option?.key as number;
                                                    const plantName = option?.text as string;
                                                    setSelectedPlantId(plantId);     // ✅ for dropdown
                                                    setSelectedPlantName(plantName);  // ✅ for saving
                                                    await loadWorkflow(plantId, SelectedDeviationType);

                                                    // ✅ FILTER PARAMETERS BASED ON PLANT
                                                    const filteredparameter: IDropdownOption[] =
                                                        ParameterTypeItems
                                                            ?.filter(item => item.PlantName === plantName)
                                                            .map(item => ({
                                                                key: item.Id,
                                                                text: item.ParameterType
                                                            }))
                                                            // remove duplicate values
                                                            .filter((item, index, self) => self.findIndex(t => t.text === item.text) === index) || [];
                                                    setParameterTypeOptions(filteredparameter);

                                                    // ✅ RESET PARAMETER & SUBPARAMETER ON PLANT CHANGE
                                                    setRows(prev => prev.map(r => ({
                                                        ...r,
                                                        parameters: undefined,
                                                        parameterText: "",
                                                        subParameter: undefined,
                                                        subParameterText: "",
                                                        subParameterOptions: [],
                                                        showOthers: false,
                                                        others: ""
                                                    })));
                                                }}
                                            />
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="Production QTY" className='font'>Production QTY <span className="Mantorystar">*</span></label>
                                            <input type="text" className='form-control' onKeyPress={(e) => { if (!/[0-9]/.test(e.key)) { e.preventDefault(); } }} onChange={e => setProductionQTY(e.target.value)} />
                                        </div>
                                    </div>
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label htmlFor="Requester Name" className='font'>Requester Name <span className="Mantorystar">*</span></label>
                                            <input type="text" value={currentUserRef.current?.Title} readOnly className='form-control readonly' />
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="ProductionDate" className='font'>Production Date <span className="Mantorystar">*</span></label>
                                            <input type="date" className='form-control' onChange={e => setProductionDate(e.target.value)} />
                                        </div>
                                        <div className='col-md-4'>
                                            <label htmlFor="Deviation Type" className='font'>Deviation Type <span className="Mantorystar">*</span></label>
                                            <Dropdown
                                                options={NPDProductionOptions}
                                                selectedKey={SelectedDeviationType}
                                                onChange={async (e, option) => {
                                                    const value = option?.key as number;
                                                    setSelectedDeviationType(value);
                                                    await loadWorkflow(selectedPlantId, value);
                                                }}
                                            />
                                        </div>
                                    </div>
                                </div>
                                <div className="heading1" style={{ marginTop: "10px" }}>
                                    <label>Upload Documents</label>
                                </div>
                                <div className="main-formcontainer">
                                    <div className='row mb-20'>
                                        <div className='col-md-4'>
                                            <label className='font'>Attachment <span className='Mantorystar'>*</span></label>
                                            <Upload
                                                beforeUpload={(file) => { setDeviationAttachFile(file); return false; }}
                                                onRemove={() => { setDeviationAttachFile(null); }} maxCount={1}
                                                className="upload-full-width">
                                                <Button className="upload-btn-full" icon={<UploadOutlined />} iconPosition="end"></Button>
                                            </Upload>
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
                                                {rows.map((row, index) => (
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
                                                                className="formtext-control"
                                                                onChange={(e, option) => {
                                                                    const selectedId = option?.key as number;
                                                                    const selectedText = option?.text as string;

                                                                    const filteredSubTypes: IDropdownOption[] =
                                                                        ParameterTypeItems
                                                                            ?.filter(
                                                                                item =>
                                                                                    item.PlantName === selectedPlantName &&
                                                                                    item.ParameterType === selectedText &&
                                                                                    item.SubParameter
                                                                            )
                                                                            .map(item => ({
                                                                                key: item.Id,
                                                                                text: item.SubParameter as string
                                                                            })) || [];

                                                                    const updatedRows = [...rows];

                                                                    updatedRows[index].parameters = selectedId;
                                                                    updatedRows[index].parameterText = selectedText; // ✅ IMPORTANT

                                                                    updatedRows[index].subParameter = undefined;
                                                                    updatedRows[index].subParameterText = "";

                                                                    updatedRows[index].subParameterOptions = filteredSubTypes;

                                                                    setRows(updatedRows);
                                                                }}
                                                            />
                                                        </td>
                                                        <td>
                                                            <Dropdown
                                                                className="formtext-control"
                                                                options={(row.subParameterOptions || []).sort((a, b) => a.text.localeCompare(b.text))}
                                                                selectedKey={row.subParameter}
                                                                // onChange={(e, option) => {
                                                                //     handleChange(index, "subParameter", option?.key);
                                                                // }}
                                                                onChange={(e, option) => {
                                                                    const updatedRows = [...rows];
                                                                    const selectedText = option?.text as string;

                                                                    updatedRows[index].subParameter = option?.key as number;        // UI
                                                                    updatedRows[index].subParameterText = option?.text as string;
                                                                    updatedRows[index].showOthers = selectedText === 'Others';
                                                                    if (selectedText !== 'Others') {
                                                                        updatedRows[index].others = "";
                                                                    }   // ✅ store TEXT

                                                                    setRows(updatedRows);
                                                                }}
                                                            />
                                                        </td>
                                                        {showOthersColumn && (
                                                            <td>
                                                                {row.showOthers ? (
                                                                    <textarea
                                                                        className='form-control'
                                                                        style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }}
                                                                        value={row.others}
                                                                        onChange={(e) => 
                                                                            handleChange(index, 'others', e.currentTarget.value)
                                                                        }
                                                                    />
                                                                ) : (
                                                                    <span>-</span>
                                                                )}
                                                            </td>
                                                        )}
                                                        <td>
                                                            <textarea name="" id="" className='form-control' style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }} value={row.specification}
                                                                onChange={(e) => {
                                                                    const textarea = e.currentTarget;
                                                                    textarea.style.height = "auto";
                                                                    textarea.style.height = textarea.scrollHeight + "px";
                                                                    handleChange(index, "specification", textarea.value);
                                                                }}
                                                            ></textarea>
                                                        </td>
                                                        <td>
                                                            <textarea className='form-control' style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }} value={row.observation}
                                                                onChange={(e) => {
                                                                    const textarea = e.currentTarget;
                                                                    textarea.style.height = "auto";
                                                                    textarea.style.height = textarea.scrollHeight + "px";
                                                                    handleChange(index, "observation", textarea.value);
                                                                }}
                                                            ></textarea>

                                                        </td>

                                                        <td>
                                                            <TextField type="date" value={row.date} onChange={(e) => handleChange(index, "date", e.currentTarget.value)} style={{ fontSize: "13px" }} />
                                                        </td>

                                                        <td>
                                                            <TextField value={row.quantity} onKeyPress={(e) => { if (!/[0-9]/.test(e.key)) { e.preventDefault(); } }} onChange={(e) => handleChange(index, "quantity", e.currentTarget.value)} style={{ fontSize: "13px" }} />
                                                        </td>

                                                        <td>
                                                            <textarea className='form-control' style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }} value={row.rootCause}
                                                                onChange={(e) => {
                                                                    const textarea = e.currentTarget;
                                                                    textarea.style.height = "auto";
                                                                    textarea.style.height = textarea.scrollHeight + "px";
                                                                    handleChange(index, "rootCause", textarea.value);
                                                                }}
                                                            ></textarea>
                                                        </td>

                                                        <td>
                                                            <textarea className='form-control' style={{ resize: "none", overflowY: "hidden", height: "auto", fontSize: "13px" }} value={row.correctiveAction}
                                                                onChange={(e) => {
                                                                    const textarea = e.currentTarget;
                                                                    textarea.style.height = "auto";
                                                                    textarea.style.height = textarea.scrollHeight + "px";
                                                                    handleChange(index, "correctiveAction", textarea.value);
                                                                }}></textarea>
                                                        </td>
                                                        <td>
                                                            <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: "5px" }}>
                                                                <DefaultButton onClick={addRow} iconProps={{ iconName: 'Add' }} className="Add" />
                                                                <DefaultButton onClick={() => deleteRow(index)} iconProps={{ iconName: 'Delete' }} className="Delete" />
                                                            </div>
                                                        </td>

                                                    </tr>
                                                ))}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>

                                <div className='row my-3'>
                                    <div className='col-md-12'>
                                        <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: "5px" }}>
                                            <DefaultButton onClick={handleSubmit} type='submit' text="Submit" className='submit-btn'
                                                value={'Save'} iconProps={{ iconName: 'Accept' }} >
                                                Submit
                                            </DefaultButton>
                                            <DefaultButton type='Cancel' text="Cancel" className='Cancel-btn'
                                                value={'Cancel'} iconProps={{ iconName: 'Cancel' }} onClick={() => history.goBack()} >
                                                Cancel
                                            </DefaultButton>
                                        </div>

                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>

        </div>
    );
};