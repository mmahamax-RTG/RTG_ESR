"use client";

import React, { useState, useEffect } from "react";
import * as XLSX from "xlsx";

const systems = [
  "DIU_1", "DIU_2", "HOIST", "SPREADER", "GANTRY", "POWER_AND_CONTROL", 
  "TROLLEY", "AUXILIARY", "BATTERY_SYSTEM", "ELECTRICAL", "GENERATING_SET", 
  "STRUCTURE", "T_ARM_1", "T_ARM_2"
];

const natureList = [
  "Mechanical Failure", "Electrical Failure", "External Influence", 
  "Material Failure", "Control System Failure"
];

const subSystemMap = {
  DIU_1: ["COLLECTOR_RAIL_SYSTEM", "COMPRESSOR_SYSTEM", "CONTROL_SYSTEM", "CURRENT_COLLECTOR_SYSTEM", "E_CYLINDER", "LOCKING_SERVO", "POSITIONING_SYSTEM"],
  DIU_2: ["COLLECTOR_RAIL_SYSTEM", "COMPRESSOR_SYSTEM", "CONTROL_SYSTEM", "CURRENT_COLLECTOR_SYSTEM", "E_CYLINDER", "LOCKING_SERVO", "POSITIONING_SYSTEM"],  
  AUXILIARY: ["AIR_CONDITIONING&VENTILATION_SYSTEM", "LCMS_SYSTEM", "LIFTING_WINCH", "LIGHTING_SYSTEM", "PA&INTERCOM_SYSTEM", "UPS", "VMT&RADIO"],
  GANTRY: ["BRAKE_SYSTEM", "GANTRY_WARNING_SYSTEM", "GANTRY_WHEEL_SYSTEM", "GEARBOX_SYSTEM", "LIMIT_INTERLOCK_SYSTEM", "MOTORS", "WHEEL_TURN_SYSTEM"],
  GENERATING_SET: ["AC_ALTERNATOR", "CONTROL_PANEL", "NORMAL_ENGINE"],
  HOIST: ["ANTI_SWAY_SYSTEM", "BRAKE_SYSTEM", "GEARBOX_SYSTEM", "HEAD_BLOCK_SYSTEM", "LIMIT_INTERLOCK_SYSTEM", "LOAD_DETECTION_SYSTEM", "MOTORS", "ROPE_REEVING_SYSTEM", "TRIM_LIST_SKEW_SYSTEM", "EMERGENCY_BRAKE_SYSTEM"],
  POWER_AND_CONTROL: ["CABLING_SYSTEM", "DRIVE_SYSTEM", "EMERGENCY_STOP_SYSTEM", "LV_DISTRIBUTION", "LV_SWITCHGEAR", "LV_TRANSFORMER", "OPERATOR_CABIN_CONSOLES", "PLC_SYSTEM", "EROOM_UPS", "POWER_CONVERSION"],
  SPREADER: ["FLIPPER_SYSTEM", "HYDRAULIC_SYSTEM", "LANDED_SYSTEM", "POWER_AND_CONTROL", "STRUCTURE_SYSTEM", "TELESCOPIC_SYSTEM", "TWIN_SYSTEM", "TWIST_LOCK_SYSTEM"],
  STRUCTURE: ["ELECTRICAL_ROOM", "EQUALIZER&BOOGIE", "GIRDERS", "MAIN_LEGS", "OPERATOR_CABIN", "SILL_BEAMS", "TROLLEY_PLATFORM", "WALKWAYS_LADDERS&PLATFORMS"],
  TROLLEY: ["BRAKE_SYSTEM", "ENERGY_CHAIN_SYSTEM", "FESTOON_SYSTEM", "GEARBOX_SYSTEM", "LIMIT_INTERLOCK_SYSTEM", "MOTORS", "TROLLEY_WHEEL_SYSTEM"],
  ELECTRICAL: ["BUSFAULT/PROFIBUS/PROFINET/FIBER OPTIC", "CONTACTOR", "FUSE", "MAIN SWITCH", "OPTIC CONVERTER", "PHASE SEQUENCE RELAY", "PLC COMPONENT", "POWER SUPPLY AC/DC", "PROGRAMME ERROR/MISSING"],
  T_ARM_1: ["BALL CASTER/GUIDE ROLLER", "BEARING", "BOLT/NUT", "BOTTOM PRESSURE ROLLER", "BRACKET SENSOR/MAGNET", "BUFFER", "CABLE CLAMP", "CABLE/TERMINAL", "CHAIN", "CONTROL/MANUAL PANEL", "COPPER SHOE - EARTH", "COPPER SHOE - PHASE", "ENERGY CHAIN", "FUSE", "INSULATOR COVER", "IO LINK/A-SI", "LASER DISTANCE/POSITION/ENCODER", "MOTOR", "MOTOR COVER", "OPTIC CONVERTER", "PANEL/JUNCTION BOX", "PE CURRENT COLLECTOR", "PH CURRENT COLLECTOR", "RELAY/CONTACTOR/CB", "SAFETY PIN", "SELECTOR/SWITCH", "SENSOR", "T ARM TELESCOPIC", "TELESCOPIC STRUCTURE", "TOP PRESSURE ROLLER", "TOUCH PANEL/HMI", "TROLLEY FRAME"],
  T_ARM_2: ["BALL CASTER/GUIDE ROLLER", "BEARING", "BOLT/NUT", "BOTTOM PRESSURE ROLLER", "BRACKET SENSOR/MAGNET", "BUFFER", "CABLE CLAMP", "CABLE/TERMINAL", "CHAIN", "CONTROL/MANUAL PANEL", "COPPER SHOE - EARTH", "COPPER SHOE - PHASE", "ENERGY CHAIN", "FUSE", "INSULATOR COVER", "IO LINK/A-SI", "LASER DISTANCE/POSITION/ENCODER", "MOTOR", "MOTOR COVER", "OPTIC CONVERTER", "PANEL/JUNCTION BOX", "PE CURRENT COLLECTOR", "PH CURRENT COLLECTOR", "RELAY/CONTACTOR/CB", "SAFETY PIN", "SELECTOR/SWITCH", "SENSOR", "T ARM TELESCOPIC", "TELESCOPIC STRUCTURE", "TOP PRESSURE ROLLER", "TOUCH PANEL/HMI", "TROLLEY FRAME"],
  BATTERY_SYSTEM: ["UPS_CONTROL", "BATTERY_CELL_MODULE", "CABLING_SYSTEM", "DRIVE_SYSTEM", "EMERGENCY_STOP_SYSTEM", "FIRE_ALARM_SYSTEM", "LV_TRANSFORMER", "PLC_SYSTEM", "BMS", "BATTERY COMM"]
};

const staffData = {
  A: [
    { name: "ZAMRI ABU BAKAR", group: "BDN-A1" },
    { name: "MOHD RUL FAUZIE JAMAL", group: "BDN-A1" },
    { name: "NURAZYZUL AHMAD", group: "BDN-A2" },
    { name: "MOHD ARIEF MD ZAILANI", group: "BDN-A2" },
    { name: "MUHAMMAD FAHMI NGAH", group: "BDN-A3" },
    { name: "MAHADIR KASDI", group: "BDN-A3" },
    { name: "MOHAMAD AMIRUL MOHD YAHYA", group: "BDN-A4" },
    { name: "MUHAMAD SAFAWI HASIM", group: "BDN-A4" },
    { name: "FARIS AZMAN", group: "UCM-A1" },
    { name: "ZAIHIN MOHD ZAHIB", group: "UCM-A1" },
    { name: "MOHD IRFAN HAFIZ DZULKIFLI", group: "UCM-A2" },
    { name: "MOHAMAD RASHIDI RAHAMAN", group: "UCM-A2" },
    { name: "HELMI HJ. BASIR", group: "UCM-A3" },
    { name: "MOHD HANAFI MOHD FOHAD", group: "UCM-A3" },
    { name: "MOHAMAD DEFRI MAHYUNIS", group: "UCM-A3" },
    { name: "MOHD FAREEZ SUITO", group: "UCM-A4" },
    { name: "MUHAMMAD AZHAR HAZEMI", group: "UCM-A4" },
    { name: "MOHD HALIM FADZILAH", group: "UCM-A4" },
    { name: "MOHAMAD FADIL TASMAN", group: "UCM-A5" },
    { name: "AMERRUL HAKIM ROSLAN", group: "UCM-A5" },
    { name: "SAYDATULNIZAM (ZOMBIE)", group: "UCM-A5" },
    { name: "AZROL AFFENDY (AWOK)", group: "E&E TEAM A" },
    { name: "ARI", group: "E&E TEAM A" },
    { name: "MOHD AZRUL FARHAN BAHARUDIN (PAAN)", group: "E&E TEAM A" },
    { name: "NOR HALIMI MISWAN", group: "E&E TEAM A" },
    { name: "MOHD NOR SYAFIQ ABD KARIM", group: "E&E TEAM A" },
    { name: "MOHD ZAKUWAN (AWE)", group: "E&E TEAM A" }
  ],
  B: [
    { name: "MOHAMAD HAIRUL AZWAN BIN SULONG", group: "B1" },
    { name: "MUHAMMED FAISAL BIN AB. KAHAR", group: "B1" },
    { name: "ABDUL RAZAK BIN ABDUL RAHMAN", group: "B2" },
    { name: "MUHAMMAD AIZUDDEN BIN MD YAACOB", group: "B2" },
    { name: "SETHURAMAN A/L VELU", group: "B3" },
    { name: "ABDUL ZULKARNAIN BIN SAMSUL", group: "B3" },
    { name: "HAIRUL ASMAWI BIN AMIN", group: "B4" },
    { name: "MUHAMMAD ASYROF KAILANI BIN ABD. RAZAK", group: "B4" },
    { name: "MOHD SUKUR BIN HANA", group: "UCM-B1" },
    { name: "MUHAMAD HAZWAN BIN MOHD ZAHRI", group: "UCM-B1" },
    { name: "KHIRUL HAFISHAM BIN KAMAL", group: "UCM-B2" },
    { name: "MUHAMMAD FAHMI BIN MOHD RAZALI", group: "UCM-B2" },
    { name: "ABDUL RAHIM BIN HASSAN", group: "UCM-B4" },
    { name: "ZAINUL ARIFFIN MOHD ZAIN", group: "UCM-B4" },
    { name: "MUHAMMAD ADIB BIN NORIZAN", group: "UCM-B3" },
    { name: "MOHAMAD SHAH REDHA BIN AHMAD TARMIZI", group: "UCM-B3" },
    { name: "SYAHRILNIZAT BIN MOHD ISA", group: "UCM-B3" },
    { name: "MOHD KAHAR BIN RAHMAT", group: "UCM-B3" },
    { name: "NADZRI MUSA", group: "CM-B" },
    { name: "HAIRIZAM WARIS", group: "CM-B" },
    { name: "HELMI HUZAIRI BIN ABDUL RAHMAN", group: "CM-B" },
    { name: "MOHAMAD NAIM BIN KAMARI", group: "E&E TEAM B" },
    { name: "SYAFIQ IZUDDIN BIN JOHARI", group: "E&E TEAM B" }
  ],
  C: [
    { name: "MOHD NOR AIMAN BIN ZAINUDDIN", group: "C2" },
    { name: "MOHD SAHAIRIL BIN KADIR", group: "C2" },
    { name: "MOHD FAZRUL AMIRUL BIN MOHAMMAD FUZI", group: "C3" },
    { name: "MOHD AFIS BIN MOHD ARIF", group: "C3" },
    { name: "MOHD IZWAN SYAH BIN SULIMAN", group: "C5" },
    { name: "ISMAIL BIN ALI", group: "C5" },
    { name: "MOHD HASSAN BIN SHAHUL HAMEED", group: "C6" },
    { name: "MOHD ZULFAEZI BIN AHMAD", group: "C6" },
    { name: "MOHD IRWAN BIN A GHANI", group: "UCM-C1" },
    { name: "MOHAMAD ALIFF BIN SAPARUDIN", group: "UCM-C1" },
    { name: "ABDUL RAHMAN BIN OYUB", group: "UCM-C2" },
    { name: "AZLAN BIN SURADI", group: "UCM-C2" },
    { name: "ABD RAZAK BIN SAMSUDIN", group: "UCM-C2" },
    { name: "SAHARIL HAMIZI BIN MOHD ISA", group: "UCM-C3" },
    { name: "AMRAN OMAR", group: "UCM-C3" },
    { name: "MOHD IKHMAL HAKIM BIN SHARIF", group: "UCM-C4" },
    { name: "ROMA ANAK KWANG", group: "UCM-C4" },
    { name: "HAIDEE BIN ANUAR", group: "CM-C1" },
    { name: "MUHAMMAD SYAFIQ BIN MUKHATAR", group: "CM-C1" },
    { name: "MUHAMMAD RIZAL IMAN BIN ABDUL HALIM", group: "E&E TEAM C" },
    { name: "MUHAMMAD AIDIL BIN OSMIRA", group: "E&E TEAM C" }
  ],
  CONTRACTOR: [
    { name: "EJSB", group: "CONTRACTOR" },
    { name: "GM", group: "CONTRACTOR" },
    { name: "KVC", group: "CONTRACTOR" },
    { name: "ZPMC", group: "CONTRACTOR" },
    { name: "GROWMATE", group: "CONTRACTOR" },
    { name: "ITP", group: "CONTRACTOR" },
    { name: "BUBENZER", group: "CONTRACTOR" },
    { name: "PTIS", group: "CONTRACTOR" },
    { name: "METROLIFT", group: "CONTRACTOR" },
    { name: "HHEP", group: "CONTRACTOR" },
    { name: "IES", group: "CONTRACTOR" }
  ]
};

export default function ESRApp() {
  const [role, setRole] = useState<"manager" | "technician" | null>(null);
  const [password, setPassword] = useState("");

  const emptyForm = {
    eqId: "",
    team: "",
    system: "",
    subSystem: "",
    fault: "",
    finding: "",
    solution: "",
    technician: [] as string[],
    group: "", 
    nature: "",
    status: "Closed",
    start: "",
    repeated: "",
    bypass: "",
    details: "",
    faultCode: "",
    operatorId: "",
    location: "",
    end: ""
  };

  const [form, setForm] = useState(emptyForm);
  const [records, setRecords] = useState<any[]>(() => {
    if (typeof window !== "undefined") {
      const saved = localStorage.getItem("esrRecords");
      return saved ? JSON.parse(saved) : [];
    }
    return [];
  });

  const [editIndex, setEditIndex] = useState<number | null>(null);

  useEffect(() => {
    localStorage.setItem("esrRecords", JSON.stringify(records));
  }, [records]);

  const totalDowntime = records.reduce((sum: number, r: any) => {
    return sum + Number(r.duration || 0);
  }, 0);

  const handleLogin = () => {
    if (password === "manager123") {
      setRole("manager");
    } else {
      setRole("technician");
    }
  };

  const handleChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement>) => {
    const { name, value } = e.target;

    if (name === "system") {
      let newFinding = "";
      if (value === "SPREADER") {
        newFinding = `Last spreader/RTG movement or activities (e.g Change Block): \nSpreader mode (twin/20/40): \nDetail of parts failure: \nDetails location of parts (e.g Valve): \nWhat action taken? What parts replace?: \nWhich location (mention RTG SLS/LSR): \nDetails workflow from first till end:`;
      }

      setForm(prev => ({
        ...prev,
        system: value,
        subSystem: "",
        fault: "",
        finding: newFinding,
        solution: ""
      }));
    } else if (name === "team") {
      setForm(prev => ({
        ...prev,
        team: value,
        technician: [],
        group: ""
      }));
    } else {
      let newValue = value;

      if (["faultCode", "operatorId", "location"].includes(name)) {
        newValue = value.toUpperCase();
      }

      if (name === "eqId") {
        newValue = value.toUpperCase();
        if (newValue && !newValue.startsWith("R")) {
          newValue = "R" + newValue;
        }
      }

      setForm(prev => ({
        ...prev,
        [name]: newValue
      }));
    }
  };

  const clearForm = () => {
    setForm(emptyForm);
  };

  const handleEdit = (index: number) => {
    const selected = records[index];
    const helperExtract = (str: string, label: string) => {
      if (!str || !str.includes(label)) return "";
      const parts = str.split(label);
      return parts.length > 1 ? parts[1].split("\n")[0].trim() : "";
    };

    const workDesc = selected.full?.["Work Description"] || "";

    setForm({
      eqId: selected.full?.["Eq ID"] || "",
      team: selected.full?.["Attend By"] ? selected.full?.["Attend By"].split(" / ")[0] : "",
      system: selected.full?.["Breakdown System"] || "",
      subSystem: selected.full?.["Sub System"] || "",
      fault: selected.fault || "",
      finding: helperExtract(workDesc, "Finding: "),
      solution: helperExtract(workDesc, "Solution: "),
      repeated: helperExtract(workDesc, "Repeated: "),
      bypass: helperExtract(workDesc, "Bypass: "),
      details: helperExtract(workDesc, "Detail: "),
      faultCode: selected.full?.["Fault Code"] || "",
      operatorId: selected.full?.["Operator ID"] || "",
      location: selected.location || "",
      start: selected.start,
      end: selected.end,
      technician: selected.tech || [],
      group: selected.full?.["Attend By"] || "",
      nature: selected.full?.["Nature Of Fault"] || "",
      status: selected.full?.["Status"] || "Closed"
    });

    setEditIndex(index);
  };

  const addRecord = () => {
    let downtime = "";

    if (form.start && form.end) {
      const start = new Date(form.start);
      const end = new Date(form.end);
      const diff = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
      downtime = diff < 0 ? "0" : diff.toFixed(2);
    }
    
    const totalTech = form.technician.length;
    const newRecord = {
      rtg: form.eqId || "-",
      fault: form.fault || "-",
      start: form.start || "-",
      end: form.end || "-",
      location: form.location || "-",
      tech: form.technician.length ? form.technician : ["-"],
      duration: downtime || "0",

      full: {
        "Eq ID": form.eqId || "-",
        "Breakdown System": form.system || "-",
        "Sub System": form.subSystem || "-",
        "Work Description":
          `Fault: ${form.fault || "-"}\n` +
          `Finding: ${form.finding || "-"}\n` +
          `Solution: ${form.solution || "-"}\n` +
          `Repeated: ${form.repeated || "-"}\n` +
          `Bypass: ${form.bypass || "-"}\n` +
          `Detail: ${form.details || "-"}`,
        "Type": "OUT",
        "Nature Of Fault": form.nature || "-",
        "Status": form.status,
        "Actual Start Date/Time": form.start || "-",
        "Complete Date/Time": form.end || "-",
        "Downtime (Hrs)": downtime || "0",
        "Work Order No.": "",
        "Attend By": form.group || "-",
        "Fault Code": form.faultCode || "-",
        "Operator ID": form.operatorId || "-",
        "Location": form.location || "-",
        "Total Technician": totalTech || 0
      }
    };

    if (editIndex !== null) {
      const updated = [...records];
      updated[editIndex] = newRecord;
      setRecords(updated);
      setEditIndex(null);
    } else {
      setRecords([...records, newRecord]);
    }
    clearForm();
  };

  const exportExcel = () => {
    if (role !== "manager") {
      alert("Only Duty Manager can export Excel");
      return;
    }

    const data = records.map((r: any) => r.full);
    const worksheet = XLSX.utils.json_to_sheet(data);

    worksheet['!cols'] = [
      { wch: 10 },{ wch: 20 },{ wch: 25 },{ wch: 50 },
      { wch: 20 },{ wch: 12 },{ wch: 20 },{ wch: 20 },
      { wch: 15 },{ wch: 15 }
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "ESR");
    XLSX.writeFile(workbook, "ESR_Report.xlsx");
  };

  if (!role) {
    return (
      <div 
        className="min-h-screen flex items-center justify-center p-4 bg-cover bg-center bg-no-repeat"
        style={{ backgroundImage: "url('/RTGFu.jpeg')" }}
      >
        <div className="absolute inset-0 bg-black/40"></div>
        <div className="relative bg-white/95 backdrop-blur p-6 rounded-xl shadow-2xl w-full max-w-sm border border-white/20">
          <h2 className="text-xl font-bold mb-4 text-center text-gray-800">Login</h2>
          <input
            type="password"
            placeholder="Enter Password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            className="w-full mb-4 p-2.5 border border-gray-300 rounded focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          <button
            onClick={handleLogin}
            className="bg-blue-600 hover:bg-blue-700 transition text-white w-full p-2.5 rounded font-semibold shadow"
          >
            Login
          </button>
        </div>
      </div>
    );
  }

  return (
    <div 
      className="min-h-screen p-2 sm:p-4 md:p-6 bg-cover bg-center bg-no-repeat bg-fixed"
      style={{ backgroundImage: "url('/RTGFu.jpeg')" }}
    >
      <div className="max-w-3xl mx-auto bg-white/95 backdrop-blur-sm p-3 sm:p-6 rounded-xl shadow-2xl border border-white/20">
        <h1 className="text-lg sm:text-xl font-bold mb-4 text-gray-800 text-center sm:text-left tracking-wide">
          ESR RTG SERVICES
        </h1>

        {/* HEADER AREA */}
        <div className="flex flex-col sm:flex-row justify-between items-stretch sm:items-center mb-4 p-2 bg-gray-100 rounded-lg gap-2">
          <div className="text-xs sm:text-sm text-gray-700 text-center sm:text-left self-center">
            Logged in as: <b>{role === "manager" ? "Duty Manager 👔" : "Technician 👷"}</b>
          </div>
          <button 
            onClick={() => { setRole(null); setPassword(""); }}
            className="bg-gray-500 hover:bg-gray-600 text-white px-4 py-1.5 rounded text-xs sm:text-sm font-medium transition text-center"
          >
            Logout
          </button>
        </div>

        {/* INPUT GRID - RESPONSIVE */}
        <div className="space-y-2">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
            <input 
              name="eqId" 
              placeholder="RTG Number (e.g. Rxx)" 
              value={form.eqId} 
              onChange={handleChange} 
              className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" 
            />
            <select
              name="team"
              value={form.team || ""}
              onChange={handleChange}
              className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500"
            >
              <option value="">Select Team</option>
              <option value="A">Team A</option>
              <option value="B">Team B</option>
              <option value="C">Team C</option>
              <option value="CONTRACTOR">Contractor</option>
            </select>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
            <select name="system" value={form.system} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500">
              <option value="">Select System</option>
              {systems.map((s, i) => <option key={i} value={s}>{s}</option>)}
            </select>

            <select name="subSystem" value={form.subSystem} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500">
              <option value="">Select Sub System</option>
              {form.system && subSystemMap[form.system as keyof typeof subSystemMap]?.map((sub, i) => (
                <option key={i} value={sub}>{sub}</option>
              ))}
            </select>
          </div>

          <textarea name="fault" placeholder="Fault e.g DIU Flt 16, Can't Trolley" value={form.fault} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" rows={2} />
          
          <textarea 
            name="finding" 
            placeholder="Finding e.g Brake no3 not released, Motor Overheat" 
            value={form.finding} 
            onChange={handleChange} 
            className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500 h-32 resize-y"
          />

          <textarea name="solution" placeholder="Solution" value={form.solution} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" rows={2} />

          <div className="grid grid-cols-1 md:grid-cols-3 gap-2">
            <input name="repeated" placeholder="Repeated? (Yes/No)" value={form.repeated} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" />
            <input name="bypass" placeholder="Bypass? (Yes/No, FC/Network #)" value={form.bypass} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" />
            <input name="details" placeholder="Any Specific Flt Code" value={form.details} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" />
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
            <input name="operatorId" placeholder="Operator ID (e.g PTP6556)" value={form.operatorId} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" />
            <input name="location" placeholder="Location (e.g 5F.58)" value={form.location} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" />
          </div>

          <select name="nature" value={form.nature} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500">
            <option value="">Nature of Fault</option>
            {natureList.map((n, i) => <option key={i} value={n}>{n}</option>)}
          </select>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
            <div>
              <label className="block text-xs font-semibold text-gray-600 mb-0.5">Start DateTime</label>
              <input type="datetime-local" name="start" value={form.start} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" />
            </div>
            <div>
              <label className="block text-xs font-semibold text-gray-600 mb-0.5">End DateTime</label>
              <input type="datetime-local" name="end" value={form.end} onChange={handleChange} className="w-full p-2 border border-gray-300 rounded text-sm bg-white focus:ring-2 focus:ring-blue-500" />
            </div>
          </div>
        </div>

        {/* TECHNICIAN LIST */}
        <div className="border border-gray-300 p-2 rounded my-3 max-h-40 overflow-y-auto bg-white shadow-inner">
          <div className="text-sm font-semibold text-gray-700">Select Technician</div>
          <div className="text-[11px] text-gray-500 mb-1.5">Tick one or more technicians below:</div>

          {form.team && staffData[form.team as keyof typeof staffData]?.map((staff, i) => (
            <label key={i} className="flex items-center text-sm py-1 border-b border-gray-50 last:border-0 cursor-pointer select-none">
              <input
                type="checkbox"
                value={staff.name}
                checked={form.technician.includes(staff.name)}
                onChange={(e) => {
                  let updated = [...form.technician];
                  if (e.target.checked) {
                    updated.push(staff.name);
                  } else {
                    updated = updated.filter(name => name !== staff.name);
                  }

                  const currentTeamStaff = staffData[form.team as keyof typeof staffData] || [];
                  const selectedStaff = currentTeamStaff.filter(s => updated.includes(s.name));
                  const groups = [...new Set(selectedStaff.filter(Boolean).map(s => s.group))];

                  setForm(prev => ({
                    ...prev,
                    technician: updated,
                    group: groups.join(" / ")
                  }));
                }}
                className="mr-2 h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
              />
              <span className="text-gray-750 text-xs sm:text-sm">{staff.name} ({staff.group})</span>
            </label>
          ))}
        </div>

        {/* CONTROL BUTTONS */}
        <div className="space-y-2 mt-4">
          <button onClick={addRecord} className="bg-green-600 hover:bg-green-700 transition text-white px-4 py-2.5 rounded w-full font-semibold shadow text-sm sm:text-base">
            {editIndex !== null ? "🔧 Update Record" : "➕ Add Record"}
          </button>

          {role === "manager" && (
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              <button
                onClick={() => {
                  if (confirm("Start new shift? All records will be cleared")) {
                    setRecords([]);
                    localStorage.removeItem("esrRecords");
                  }
                }}
                className="bg-red-600 hover:bg-red-700 transition text-white px-4 py-2 rounded w-full text-xs sm:text-sm font-medium shadow"
              >
                🧹 Clear Shift Data
              </button>
              <button onClick={exportExcel} className="bg-blue-600 hover:bg-blue-700 transition text-white px-4 py-2 rounded w-full text-xs sm:text-sm font-medium shadow">
                📊 Export All to Excel
              </button>
            </div>
          )}
        </div>

        {/* LOG RECORDS VIEW */}
        <div className="mt-6 border-t border-gray-200 pt-4">
          {role === "manager" && (
            <div className="bg-yellow-100 p-2.5 rounded mb-3 text-xs sm:text-sm text-gray-800 font-medium">
              📊 <b>Total Downtime:</b> {totalDowntime.toFixed(2)} hrs
            </div>
          )}

          <h2 className="font-bold text-base text-gray-800 mb-2">Logged Records</h2>

          <div className="space-y-2 max-h-96 overflow-y-auto pr-1">
            {records.length === 0 ? (
              <p className="text-xs text-gray-400 italic text-center py-4">No records found for current session.</p>
            ) : (
              records.map((r: any, i: number) => (
                <div key={i} className={`border p-3 rounded shadow-sm text-xs sm:text-sm transition-all ${editIndex === i ? "bg-yellow-50 border-yellow-500" : "bg-gray-50 border-gray-200"}`}>
                  <div className="flex justify-between items-start gap-2">
                    <div>
                      <div><b>RTG:</b> <span className="text-blue-700 font-semibold">{r.rtg}</span></div>
                      <div className="mt-0.5"><b>Fault:</b> {r.fault}</div>
                    </div>
                    {role === "manager" && (
                      <button onClick={() => handleEdit(i)} className="bg-yellow-500 hover:bg-yellow-600 text-white px-2.5 py-1 rounded text-xs font-medium transition">
                        Edit
                      </button>
                    )}
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-1 mt-2 text-xs text-gray-600 border-t border-gray-100 pt-1.5">
                    <div><b>Start:</b> {r.start}</div>
                    <div><b>End:</b> {r.end}</div>
                    <div className={`font-semibold ${Number(r.duration) > 0.5 ? "text-red-600" : "text-gray-700"}`}><b>Duration:</b> {r.duration} hrs</div>
                    <div><b>Location:</b> {r.location}</div>
                  </div>

                  <div className="text-xs text-gray-500 mt-1">
                    <b>Techs:</b> {Array.isArray(r.tech) ? r.tech.join(", ") : r.tech}
                  </div>
                </div>
              ))
            )}
          </div>

          {/* DASHBOARD POWER BI */}
          <div className="mt-6 border-t border-gray-200 pt-4">
            <h2 className="font-bold text-gray-800 text-sm sm:text-base mb-2">72hrs Repeated Dashboard</h2>
            <div className="w-full h-[400px] sm:h-[600px] md:h-[800px] lg:h-[1050px] rounded-xl overflow-hidden shadow-lg border border-gray-300 bg-white">
              <iframe
                title="72Hrs Repeated - Tech View"
                src="https://app.powerbi.com/view?r=eyJrIjoiZjJjMDA1NDAtZjJhNi00ZTdiLTk0MGYtODhmM2E5YTk1NzdmIiwidCI6ImJmZTA2YmQ3LTdiZWUtNGQ2MS1hMzVlLWRmZmYyMzQ2ZjFjMiIsImMiOjEwfQ%3D%3D"
                className="w-full h-full border-0"
                allowFullScreen
              ></iframe>
            </div>
          </div>
        </div>

      </div>
      
      {/* CORNER WATERMARK - ANTICIPATES BG OVERLAP */}
      <div className="fixed bottom-2 right-3 text-[10px] sm:text-xs text-white drop-shadow-[0_1px_2px_rgba(0,0,0,0.9)] opacity-60 pointer-events-none select-none font-sans font-medium tracking-wide">
        by5533 M&R © 2026
      </div>
    </div>
  );
}
