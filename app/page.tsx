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

export default function ESRApp() {

const [role, setRole] = useState(null);
const [password, setPassword] = useState("");

  const emptyForm = {
    eqId: "",
    system: "",
    subSystem: "",
    fault: "",
    finding: "",
    solution: "",
    technician: "",
    nature: "",
    status: "Closed",
    start: "",
repeated: "",
bypass: "",     
details: "",
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

const totalDowntime = records.reduce((sum, r) => {
  return sum + parseFloat(r.duration || 0);
}, 0);

// ✅ LOGIN FUNCTION LETAK SINI
  const handleLogin = () => {
    if (password === "manager123") {
      setRole("manager");
    } else {
      setRole("technician");
    }
  };

  const handleChange = (e) => {
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

  const addRecord = () => {
    let downtime = "";

if (form.start && form.end) {
      const start = new Date(form.start);
      const end = new Date(form.end);
      const diff = (end - start) / (1000 * 60 * 60);
      downtime = diff.toFixed(2);
    }

    const newRecord = {
      rtg: form.eqId || "-",
      fault: form.fault || "-",
      start: form.start || "-",
      end: form.end || "-",
      tech: form.technician || "-",
duration: downtime || "0",

      // full data for excel
      full: {
        "Eq ID": form.eqId || "-",
        "Breakdown System": form.system || "-",
        "Sub System": form.subSystem || "-",
        "Work Description": `Fault: ${form.fault || "-"}\nFinding: ${form.finding || "-"}\nSolution: ${form.solution || "-"}\nRepeated: ${form.repeated || "-"}\nBypass: ${form.bypass || "-"}\nDetail: ${form.details || "-"}`,
        "Nature Of Fault": form.nature || "-",
        "Status": form.status,
        "Start Time": form.start || "-",
        "End Time": form.end || "-",
        "Downtime (Hrs)": downtime || "0",
        "Attend By": form.technician || "-"
      }
    };

    setRecords([...records, newRecord]);
    clearForm(); // auto clear after add ✅
  };

  const exportExcel = () => {

  // ✅ BONUS SECURITY
  if (role !== "manager") {
    alert("Only Duty Manager can export Excel");
    return;
  }

  const data = records.map(r => r.full);
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
// ✅ ✅ LETAK LOGIN SCREEN BLOCK DI SINI
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

        <select name="system" value={form.system} onChange={handleChange} className="w-full mb-2 p-2 border rounded">
          <option value="">Select System</option>
          {systems.map((s, i) => <option key={i}>{s}</option>)}
        </select>

        <select name="subSystem" value={form.subSystem} onChange={handleChange} className="w-full mb-2 p-2 border rounded">
          <option value="">Select Sub System</option>
          {form.system && subSystemMap[form.system]?.map((sub, i) => (
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
<textarea name="bypass" placeholder="Bypass type Yes/No" value={form.bypass} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />
<textarea name="details" placeholder="Any Flt Code e.g F005, A0501, F07900 " value={form.details} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />

        <select name="nature" value={form.nature} onChange={handleChange} className="w-full mb-2 p-2 border rounded">
          <option value="">Nature of Fault</option>
          {natureList.map((n, i) => <option key={i}>{n}</option>)}
        </select>

        <input type="datetime-local" name="start" value={form.start} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />
        <input type="datetime-local" name="end" value={form.end} onChange={handleChange} className="w-full mb-2 p-2 border rounded" />

        <input name="technician" placeholder="Technician Name" value={form.technician} onChange={handleChange} className="w-full mb-4 p-2 border rounded" />

        <button onClick={addRecord} className="bg-green-600 text-white px-4 py-2 rounded w-full mb-2">
          Add Record
        </button>

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
            <div key={i} className="border p-3 mb-2 rounded bg-gray-50">
              <div><b>RTG:</b> {r.rtg}</div>
              <div><b>Fault:</b> {r.fault}</div>
              <div><b>Start:</b> {r.start}</div>
              <div><b>End:</b> {r.end}</div>
<div><b>Duration:</b> {r.duration} hrs</div>
              <div><b>Technician:</b> {r.tech}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}
