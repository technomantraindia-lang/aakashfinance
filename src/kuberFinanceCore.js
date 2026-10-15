const users = [];
const clients = [];
const vehicles = [];
const dueTasks = [];
const listings = [];
const importRows = [];
const verificationItems = [];
const callerOutcomes = [
  { outcome: "Connected - Will Pay", required: "Promise date and expected amount", nextAction: "Follow up on commitment date" },
  { outcome: "Already Paid", required: "Receipt or transaction reference", nextAction: "Pause after submission" },
  { outcome: "Insurance Interested", required: "Expected renewal date and notes", nextAction: "Move to renewal follow-up" },
  { outcome: "No Answer / Busy", required: "Attempt note", nextAction: "Schedule next attempt" },
  { outcome: "Switched Off", required: "Attempt note", nextAction: "Schedule retry" },
  { outcome: "Callback Requested", required: "Callback date, time and note", nextAction: "Schedule callback" },
  { outcome: "Wrong Number", required: "Correction note", nextAction: "Escalate to Admin" },
  { outcome: "Dispute", required: "Detailed notes", nextAction: "Escalate to Admin" }
];
const marketplaceChats = [
  { listingId: "m1", buyer: "Shree Cargo Movers", seller: "Rajesh Logistics", status: "Negotiating", unread: 2 },
  { listingId: "m2", buyer: "Rajesh Logistics", seller: "Shree Cargo Movers", status: "Interested", unread: 0 }
];
const permissionRows = [
  ["View all clients", "Yes", "No", "Assigned only", "No"],
  ["View own fleet", "Yes", "Yes", "Assigned only", "Yes"],
  ["Import Excel", "Yes", "No", "No", "No"],
  ["Edit closing principal", "Yes", "No", "No", "No"],
  ["Mark EMI paid", "Yes", "Own fleet", "No", "Own fleet"],
  ["Verify payment", "Yes", "No", "No", "No"],
  ["Call / WhatsApp", "Optional", "No", "Yes", "No"],
  ["Create sale listing", "Yes", "Yes", "No", "Yes"],
  ["Approve listing", "Yes", "No", "No", "No"],
  ["Owner chat", "Reports only", "Yes", "No", "Yes"]
];
const STORAGE_KEY = "kuber-finance-admin-data-v1";
function getNavSections(role) {
  if (role === "Admin") {
    return [
      ["dashboard", "Overview", "home-outline"],
      ["clients", "Clients", "people-outline"],
      ["fleet", "Fleet", "bus-outline"],
      ["dues", "Dues", "calendar-outline"],
      ["verification", "Verification", "checkmark-done-outline"],
      ["marketplace", "Marketplace", "storefront-outline"],
      ["reports", "Reports", "bar-chart-outline"],
      ["settings", "Settings", "settings-outline"]
    ];
  }
  if (role === "Caller") {
    return [
      ["dashboard", "Overview", "home-outline"],
      ["caller", "Queue", "call-outline"],
      ["dues", "Dues", "calendar-outline"],
      ["clients", "Clients", "people-outline"],
      ["profile", "Profile", "person-outline"]
    ];
  }
  return [
    ["dashboard", "Overview", "home-outline"],
    ["fleet", "Fleet", "bus-outline"],
    ["dues", "Dues", "calendar-outline"],
    ["marketplace", "Marketplace", "storefront-outline"],
    ["profile", "Profile", "person-outline"]
  ];
}
const initialAuditLogs = [];
const initialNotifications = [];
const initialCallerActivities = [];
const initialDocuments = [];
const initialMarketplaceThreads = [];
const initialWhatsAppTemplates = [
  {
    id: "payment-reminder",
    name: "Payment reminder",
    body: "Hello {{customer}}, Kuber Finance reminder: {{vehicle}} has {{type}} due on {{dueDate}}. Amount {{amount}}. Please upload payment/renewal proof.",
    active: true
  },
  {
    id: "document-follow-up",
    name: "Document follow-up",
    body: "Hello {{customer}}, please share the pending {{type}} proof for {{vehicle}} so we can update your account.",
    active: true
  },
  {
    id: "promise-follow-up",
    name: "Promise follow-up",
    body: "Hello {{customer}}, following up on your {{type}} payment for {{vehicle}}. Please reply with an update.",
    active: true
  }
];
const initialWhatsAppLogs = [];
const initialAppData = {
  clients,
  vehicles,
  dueTasks,
  listings,
  verificationItems,
  importRows,
  clientImports: [],
  callerActivities: initialCallerActivities,
  auditLogs: initialAuditLogs,
  notifications: initialNotifications,
  documents: initialDocuments,
  marketplaceThreads: initialMarketplaceThreads,
  whatsappTemplates: initialWhatsAppTemplates,
  whatsappLogs: initialWhatsAppLogs,
  reminderSettings: { enabled: true, intervalHours: 24, windowDays: 30 }
};

function buildNextCycleDueTask(task, vehicle) {
  if (!task || !vehicle || task.status === "Closed") return null;
  const currentDate = new Date(`${task.dueDate}T00:00:00`);
  if (Number.isNaN(currentDate.getTime())) return null;
  let nextDate = null;
  let amount = task.amount;
  if (task.type === "EMI") {
    const nextSchedule = (vehicle.emiSchedule || [])
      .filter((entry) => entry.status !== "Paid")
      .map((entry) => ({ entry, date: new Date(`${entry.dueDate}T00:00:00`) }))
      .filter(({ date }) => !Number.isNaN(date.getTime()) && date > currentDate)
      .sort((left, right) => left.date.getTime() - right.date.getTime())[0];
    nextDate = nextSchedule?.date || new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, currentDate.getDate());
    amount = nextSchedule?.entry.amount ?? vehicle.emiAmount ?? task.amount;
    if (!nextSchedule && vehicle.tenure && (vehicle.paidEmi || 0) >= vehicle.tenure) return null;
  } else if (["Insurance", "Permit", "Fitness", "PUC", "Tax"].includes(task.type)) {
    nextDate = new Date(currentDate.getFullYear() + 1, currentDate.getMonth(), currentDate.getDate());
  }
  if (!nextDate || Number.isNaN(nextDate.getTime())) return null;
  const dueDate = `${nextDate.getFullYear()}-${String(nextDate.getMonth() + 1).padStart(2, "0")}-${String(nextDate.getDate()).padStart(2, "0")}`;
  const days = Math.round((nextDate.getTime() - new Date().setHours(0, 0, 0, 0)) / 86400000);
  const priority = days <= 7 ? "High" : days <= 15 ? "Medium" : "Low";
  const vehicleKey = String(vehicle.id).replace(/[^a-z0-9]/gi, "");
  return {
    id: `auto-${task.type.toLowerCase()}-${vehicleKey}-${dueDate}`,
    clientId: task.clientId,
    vehicleId: task.vehicleId,
    type: task.type,
    amount: Number(amount || 0),
    dueDate,
    status: "Due",
    callerId: task.callerId || "",
    priority
  };
}

function applyVerifiedCycleToVehicle(vehicle, task, verifiedAt, nextTask) {
  const next = { ...vehicle };
  if (task.type === "EMI") {
    const schedule = [...(vehicle.emiSchedule || [])];
    const current = schedule.find((entry) => entry.dueDate === task.dueDate);
    if (current) {
      current.status = "Paid";
      current.paidAt = verifiedAt;
      current.reference = current.reference || `EMI-${String(current.installment).padStart(3, "0")}`;
      next.paidEmi = Math.max(vehicle.paidEmi || 0, current.installment);
      next.emiSchedule = schedule;
      const history = [...(vehicle.emiHistory || [])];
      if (!history.some((entry) => entry.installment === current.installment)) {
        next.emiHistory = [...history, { installment: current.installment, amount: current.amount, paidOn: verifiedAt, reference: current.reference, status: "Verified" }];
      }
    } else if (vehicle.tenure) {
      next.paidEmi = Math.min(vehicle.tenure, (vehicle.paidEmi || 0) + 1);
    }
  }
  if (nextTask) {
    if (task.type === "Insurance") next.insuranceExpiry = nextTask.dueDate;
    if (task.type === "Permit") next.permitExpiry = nextTask.dueDate;
    if (task.type === "PUC") next.pucExpiry = nextTask.dueDate;
    if (task.type === "Fitness") next.fitnessExpiry = nextTask.dueDate;
    if (task.type === "Tax") next.nationalPermitExpiry = nextTask.dueDate;
  }
  next.insuranceHistory = (vehicle.insuranceHistory || []).map((entry) => entry.expiryDate === task.dueDate ? { ...entry, status: "Active", verifiedAt } : entry);
  next.complianceHistory = (vehicle.complianceHistory || []).map((entry) => entry.expiryDate === task.dueDate ? { ...entry, status: "Active", verifiedAt } : entry);
  return next;
}
const formatMoney = (value) => `INR ${new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(value)}`;
const liability = (vehicle) => vehicle.principal + vehicle.overdue + vehicle.penalty + vehicle.foreclosure;
const getClient = (clientId) => clients.find((client) => client.id === clientId);
const getVehicle = (vehicleId) => vehicles.find((vehicle) => vehicle.id === vehicleId);
export {
  STORAGE_KEY,
  applyVerifiedCycleToVehicle,
  buildNextCycleDueTask,
  callerOutcomes,
  clients,
  dueTasks,
  formatMoney,
  getClient,
  getNavSections,
  getVehicle,
  importRows,
  initialAppData,
  initialAuditLogs,
  initialCallerActivities,
  initialDocuments,
  initialMarketplaceThreads,
  initialNotifications,
  initialWhatsAppLogs,
  initialWhatsAppTemplates,
  liability,
  listings,
  marketplaceChats,
  permissionRows,
  users,
  vehicles,
  verificationItems
};
