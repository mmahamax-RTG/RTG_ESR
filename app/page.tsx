"use client";
import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

const systems = ["DIU_1","DIU_2","HOIST","SPREADER","GANTRY","POWER_AND_CONTROL","TROLLEY","AUXILIARY","BATTERY_SYSTEM","ELECTRICAL","GENERATING_SET","STRUCTURE","T_ARM_1","T_ARM_2"];
const natureList = ["Mechanical Failure","Electrical Failure","External Influence","Material Failure","Control System Failure"];

const subSystemMap = {
  DIU_1: ["COLLECTOR_RAIL_SYSTEM","COMPRESSOR_SYSTEM","CONTROL_SYSTEM","CURRENT_COLLECTOR_SYSTEM","E_CYLINDER","LOCKING_SERVO","POSITIONING_SYSTEM"],
  DIU_2: ["COLLECTOR_RAIL_SYSTEM","COMPRESSOR_SYSTEM","CONTROL_SYSTEM","CURRENT_COLLECTOR_SYSTEM","E_CYLINDER","LOCKING_SERVO","POSITIONING_SYSTEM"],  
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
  const [records, setRecords] = useState(() => {
  if (typeof window !== "undefined") {
    const saved = localStorage.getItem("esrRecords");
    return saved ? JSON.parse(saved) : [];
  }
  return [];
});
useEffect(() => {
  localStorage.setItem("esrRecords", JSON.stringify(records));
}, [records]);

const totalDowntime = records.reduce((sum: number, r: any) => {
  return sum + Number(r.duration || 0);
}, 0);

// ✅ LOGIN FUNCTION LETAK SINI
  const handleLogin = () => {
    if (password === "manager123") {
      setRole("manager");
    } else {
      setRole("technician");
    }
  };

  const handleChange = (e: any) => {
  const { name, value } = e.target;

  if (name === "system") {

    let newFinding = "";
    let newFault = "";
    let newSolution = "";

    if (value === "SPREADER") {
      newFinding = `Last spreader/RTG movement or activities (e.g Change Block): 
Spreader mode (twin/20/40): 
Detail of parts failure: 
Details location of parts (e.g Valve): 
What action taken? What parts replace?: 
Which location (mention RTG SLS/LSR): 
Details workflow from first till end:`;
    }

    setForm({
      ...form,
      system: value,
      subSystem: "",
      fault: newFault,
      finding: newFinding,
      solution: newSolution
    });

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

    setForm({
      ...form,
      [name]: newValue
    });
  }
};



  const clearForm = () => {
  setForm(emptyForm);
};


const [editIndex, setEditIndex] = useState<number | null>(null);

const handleEdit = (index: number) => {
  const selected = records[index];

  setForm({
  ...form,
  eqId: selected.full?.["Eq ID"] || "",
  system: selected.full?.["Breakdown System"] || "",
  subSystem: selected.full?.["Sub System"] || "",

  fault: selected.fault || "",

  finding: selected.full?.["Work Description"]?.split("Finding: ")[1]?.split("\n")[0] || "",
  solution: selected.full?.["Work Description"]?.split("Solution: ")[1]?.split("\n")[0] || "",
  repeated: selected.full?.["Work Description"]?.split("Repeated: ")[1]?.split("\n")[0] || "",
  bypass: selected.full?.["Work Description"]?.split("Bypass: ")[1]?.split("\n")[0] || "",
  details: selected.full?.["Work Description"]?.split("Detail: ")[1] || "",

  faultCode: selected.full?.["Fault Code"] || "",
  operatorId: selected.full?.["Operator ID"] || "",
  location: selected.location || "",

  start: selected.start,
  end: selected.end,

  technician: selected.tech,
  group: selected.full?.["Attend By"] || "",

  nature: selected.full?.["Nature Of Fault"] || "",
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

      // full data for excel
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
    <div className="min-h-screen flex items-center justify-center bg-gray-100">
      <div className="bg-white p-6 rounded shadow w-80">
        <h2 className="text-lg font-bold mb-4">Login</h2>

        <input
          type="password"
          placeholder="Enter Password"
          value={password}
          onChange={(e) => setPassword(e.target.value)}
          className="w-full mb-3 p-2 border rounded"
        />

        <button
          onClick={handleLogin}
          className="bg-blue-600 text-white w-full p-2 rounded"
        >
          Login
        </button>
      </div>
    </div>
  );
}

  return (
    <div className="min-h-screen bg-gray-100 p-6">
      <div className="max-w-3xl mx-auto bg-white p-6 rounded-xl shadow">
        <h1 className="text-xl font-bold mb-4">ESR RTG SERVICES</h1>

<div className="flex justify-between items-center mb-3">
  <div className="text-sm text-gray-600">
    Logged in as: <b>{role === "manager" ? "Duty Manager 👔" : "Technician 👷"}</b>
  </div>

  <button 
    onClick={() => setRole(null)}
    className="bg-gray-500 text-white px-3 py-1 rounded"
  >
    Logout
  </button>
</div>


        <input name="eqId" placeholder="RTG Number-Rxx-" value={form.eqId} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />
<select
  name="team"
  value={form.team || ""}
  onChange={handleChange}
  className="w-full mb-2 p-2 border rounded"
>
  <option value="">Select Team</option>
  <option value="A">Team A</option>
  <option value="B">Team B</option>
  <option value="C">Team C</option>
<option value="CONTRACTOR">Contractor</option>
</select>

        <select name="system" value={form.system} onChange={handleChange} className="w-full mb-2 p-2 border rounded">
          <option value="">Select System</option>
          {systems.map((s, i) => <option key={i}>{s}</option>)}
        </select>

        <select name="subSystem" value={form.subSystem} onChange={handleChange} className="w-full mb-2 p-2 border rounded">
          <option value="">Select Sub System</option>
          {form.system && subSystemMap[form.system as keyof typeof subSystemMap]?.map((sub, i) => (
            <option key={i} value={sub}>{sub}</option>
          ))}
        </select>

        <textarea name="fault" placeholder="Fault e.g DIU Flt 16, Can't Trolley, Can't Lock, Engine Can't Start" value={form.fault} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />
        
<textarea 
  name="finding" 
  placeholder="Finding e.g Brake no3 not released, Anstisway LSL not Roll, Motor Overheat, PLC Bussfault address 21" 
  value={form.finding} 
  onChange={handleChange} 
  className="w-full mb-2 p-2 border rounded h-40 resize-y"
/>

        <textarea name="solution" placeholder="Solution" value={form.solution} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />
<textarea name="repeated" placeholder="Repeated type Yes/No" value={form.repeated} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />
<textarea name="bypass" placeholder="Bypass type Yes/No, If yes, Insert FC and Network number" value={form.bypass} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />
<textarea name="details" placeholder="Any Flt Code e.g F005, A0501, F07900 " value={form.details} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />

{/* ✅ TAMBAH DI SINI */}
<input
  name="faultCode"
  placeholder="Fault Code (e.g F005)"
  value={form.faultCode || ""}
  onChange={handleChange}
  className="w-full mb-2 p-2 border rounded"
/>

<input
  name="operatorId"
  placeholder="Operator ID (e.g PTP6556)"
  value={form.operatorId || ""}
  onChange={handleChange}
  className="w-full mb-2 p-2 border rounded"
/>

<input
  name="location"
  placeholder="Location (e.g 5F.58, T/Plate 4L/5L)"
  value={form.location || ""}
  onChange={handleChange}
  className="w-full mb-2 p-2 border rounded"
/>

        <select name="nature" value={form.nature} onChange={handleChange} className="w-full mb-2 p-2 border rounded">
          <option value="">Nature of Fault</option>
          {natureList.map((n, i) => <option key={i}>{n}</option>)}
        </select>

        <input type="datetime-local" name="start" value={form.start} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />
        <input type="datetime-local" name="end" value={form.end} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />

        <div className="border p-2 rounded mb-4 max-h-40 overflow-y-auto">

  <div className="text-sm mb-1 font-semibold">
    Select Technician
  </div>

  <div className="text-xs text-gray-500 mb-2">
    Tick one or more technicians
  </div>

  {form.team && staffData[form.team as keyof typeof staffData]?.map((staff, i) => (
    <label key={i} className="block text-sm">
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

const selectedStaff = updated.map(name =>
  staffData[form.team].find(s => s.name === name)
);

const groups = [
  ...new Set(
    selectedStaff
      .filter(Boolean)
      .map(s => s.group)
  )
];

setForm({
  ...form,
  technician: updated,
  group: groups.join(" / ")
});
        }}
        className="mr-2"
      />
      {staff.name} ({staff.group})
    </label>
  ))}

</div>

          <button onClick={addRecord} className="bg-green-600 text-white px-4 py-2 rounded w-full mb-2">
          {editIndex !== null ? "Update Record" : "Add Record"}
        </button>

{role === "manager" && (
  <button
    onClick={() => {
      if (confirm("Start new shift? All records will be cleared")) {
        setRecords([]);
        localStorage.removeItem("esrRecords");
      }
    }}
    className="bg-red-600 text-white px-4 py-2 rounded w-full mb-2"
  >
    Start New Shift (Clear Data)
  </button>
)}


       {role === "manager" && (
  <button 
    onClick={exportExcel} 
    className="bg-blue-600 text-white px-4 py-2 rounded w-full mb-4"
  >
    Export All to Excel
  </button>
)}

<div className="mt-4">
  {role === "manager" && (
    <div className="bg-yellow-100 p-3 rounded mb-3">
      <b>Total Downtime:</b> {totalDowntime.toFixed(2)} hrs
    </div>
  )}

  <h2 className="font-bold mb-2">Records</h2>

  {records.map((r, i) => (
    <div
      key={i}
      className={`border p-3 mb-2 rounded ${
        editIndex === i ? "bg-yellow-100 border-yellow-500" : "bg-gray-50"
      }`}
    >
      <div><b>RTG:</b> {r.rtg}</div>
      <div><b>Fault:</b> {r.fault}</div>

      <div className="grid grid-cols-2 gap-2">
        <div><b>Start:</b> {r.start}</div>
        <div><b>End:</b> {r.end}</div>

        <div className={`font-semibold ${Number(r.duration) > 0.5 ? "text-red-600" : ""}`}>
          <b>Duration:</b> {r.duration} hrs
        </div>

        <div><b>Location:</b> {r.location}</div>
      </div>

      <div>
        <b>Technician:</b> {Array.isArray(r.tech) ? r.tech.join(", ") : r.tech}
      </div>

      {role === "manager" && (
        <button
          onClick={() => handleEdit(i)}
          className="mt-2 bg-yellow-500 text-white px-2 py-1 rounded"
        >
          Edit
        </button>
      )}
    </div>
 ))}
<div className="fixed bottom-2 right-3 text-[15px] text-gray-400 opacity-50 hover:opacity-100 transition">
  by5533 M&R © 2026
</div>
</div>

      </div>
    </div>
);
}


