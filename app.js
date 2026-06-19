const releaseItems = [
  {
    rmi: "RMI-10452",
    project: "Customer onboarding workflow upgrade",
    pm: "Anita Sharma",
    rm: "Rahul Mehta",
    app: "Onboarding Portal",
    planner: "Neha Gupta",
    rmiStatus: "Green",
    plannedStart: "21:00",
    plannedEnd: "21:45",
    actualStart: "21:02",
    actualEnd: "21:38",
    cr: "CRQ-77821",
    poc: "S. Kumar",
    remark: "Completed",
    deploymentStatus: "Completed"
  },
  {
    rmi: "RMI-10467",
    project: "Payment service version uplift",
    pm: "Vikram Rao",
    rm: "Rahul Mehta",
    app: "Payment Gateway",
    planner: "Neha Gupta",
    rmiStatus: "Orange",
    plannedStart: "21:45",
    plannedEnd: "22:30",
    actualStart: "21:48",
    actualEnd: "",
    cr: "CRQ-77844",
    poc: "A. Khan",
    remark: "Deployment in progress",
    deploymentStatus: "Ongoing"
  },
  {
    rmi: "RMI-10479",
    project: "Policy document service release",
    pm: "Meera Iyer",
    rm: "Rahul Mehta",
    app: "Document Hub",
    planner: "Amit Jain",
    rmiStatus: "Red",
    plannedStart: "22:30",
    plannedEnd: "23:10",
    actualStart: "",
    actualEnd: "",
    cr: "CRQ-77902",
    poc: "P. Nair",
    remark: "Waiting for blocker clearance",
    deploymentStatus: "Pending"
  },
  {
    rmi: "RMI-10503",
    project: "Analytics dashboard monthly release",
    pm: "Anita Sharma",
    rm: "Rahul Mehta",
    app: "Insights Dashboard",
    planner: "Amit Jain",
    rmiStatus: "Green",
    plannedStart: "23:10",
    plannedEnd: "23:50",
    actualStart: "",
    actualEnd: "",
    cr: "CRQ-77914",
    poc: "R. Das",
    remark: "Ready for release window",
    deploymentStatus: "Pending"
  }
];

const incidents = [
  {
    ticket: "JIRA-49211",
    app: "Document Hub",
    severity: "High",
    status: "Open",
    owner: "P. Nair"
  },
  {
    ticket: "JIRA-49234",
    app: "Payment Gateway",
    severity: "Medium",
    status: "Monitoring",
    owner: "A. Khan"
  },
  {
    ticket: "JIRA-49256",
    app: "Onboarding Portal",
    severity: "Low",
    status: "Resolved",
    owner: "S. Kumar"
  }
];

const statusClassMap = {
  Green: "green-text",
  Orange: "orange",
  Red: "red-text"
};

const viewTitles = {
  dashboard: "Dashboard",
  release: "Release",
  rmi: "RMI Tracking",
  deployment: "Deployment",
  releaseDay: "Release Day Command Center"
};

function setText(id, value) {
  document.getElementById(id).textContent = value;
}

function createCell(text) {
  const cell = document.createElement("td");
  cell.textContent = text;
  return cell;
}

function renderDashboard() {
  const completed = releaseItems.filter((item) => item.deploymentStatus === "Completed").length;
  const ongoing = releaseItems.filter((item) => item.deploymentStatus === "Ongoing").length;
  const pending = releaseItems.filter((item) => item.deploymentStatus === "Pending").length;

  setText("completed-apps", completed);
  setText("ongoing-apps", ongoing);
  setText("pending-apps", pending);
  setText("jira-count", incidents.length);

  const statusList = document.getElementById("application-status");
  statusList.innerHTML = "";

  releaseItems.forEach((item) => {
    const row = document.createElement("div");
    row.className = "app-status-row";

    const appBlock = document.createElement("div");
    const appName = document.createElement("strong");
    const detail = document.createElement("div");
    appName.textContent = item.app;
    detail.textContent = `${item.rmi} - ${item.project}`;
    detail.className = "metric-label";
    appBlock.append(appName, detail);

    const status = document.createElement("span");
    status.className = `status-dot ${item.deploymentStatus === "Completed" ? "green-text" : item.deploymentStatus === "Ongoing" ? "orange" : "red-text"}`;
    status.textContent = item.deploymentStatus;

    row.append(appBlock, status);
    statusList.append(row);
  });
}

function renderReleaseTable() {
  const body = document.getElementById("release-table");
  body.innerHTML = "";
  releaseItems.forEach((item) => {
    const row = document.createElement("tr");
    row.append(
      createCell(item.rmi),
      createCell(item.project),
      createCell(item.pm),
      createCell(item.rm),
      createCell(item.app)
    );
    body.append(row);
  });
}

function renderRmiTable() {
  const body = document.getElementById("rmi-table");
  body.innerHTML = "";
  releaseItems.forEach((item) => {
    const row = document.createElement("tr");
    const statusCell = document.createElement("td");
    const status = document.createElement("span");
    status.className = `status-dot ${statusClassMap[item.rmiStatus]}`;
    status.textContent = item.rmiStatus;
    statusCell.append(status);
    row.append(createCell(item.rmi), statusCell, createCell(item.planner));
    body.append(row);
  });
}

function renderDeploymentTable() {
  const body = document.getElementById("deployment-table");
  body.innerHTML = "";
  releaseItems.forEach((item) => {
    const row = document.createElement("tr");
    row.append(
      createCell(item.rmi),
      createCell(item.app),
      createCell(item.plannedStart),
      createCell(item.plannedEnd),
      createCell(item.actualStart || "-"),
      createCell(item.actualEnd || "-"),
      createCell(item.cr),
      createCell(item.poc),
      createCell(item.remark)
    );
    body.append(row);
  });
}

function renderIncidentTable() {
  const body = document.getElementById("incident-table");
  body.innerHTML = "";
  incidents.forEach((incident) => {
    const row = document.createElement("tr");
    row.append(
      createCell(incident.ticket),
      createCell(incident.app),
      createCell(incident.severity),
      createCell(incident.status),
      createCell(incident.owner)
    );
    body.append(row);
  });
}

function setupNavigation() {
  document.querySelectorAll(".nav-item").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".nav-item").forEach((item) => item.classList.remove("active"));
      document.querySelectorAll(".view").forEach((view) => view.classList.remove("active"));
      button.classList.add("active");
      document.getElementById(button.dataset.view).classList.add("active");
      setText("view-title", viewTitles[button.dataset.view]);
    });
  });
}

function setupReleaseControls() {
  const releaseState = document.getElementById("release-state");
  document.getElementById("start-release").addEventListener("click", () => {
    releaseState.textContent = "Release Started";
    releaseState.style.background = "#dcfce7";
    releaseState.style.color = "#16825d";
  });

  document.getElementById("stop-release").addEventListener("click", () => {
    releaseState.textContent = "Release Stopped";
    releaseState.style.background = "#fee2e2";
    releaseState.style.color = "#c2414b";
  });
}

function setupOverallStatus() {
  document.querySelectorAll(".status").forEach((button) => {
    button.addEventListener("click", () => {
      const status = button.dataset.status;
      setText("current-overall-status", status);
      setText("overall-pill", status);
      document.getElementById("overall-pill").className = `pill ${status.toLowerCase()}`;
    });
  });
}

renderDashboard();
renderReleaseTable();
renderRmiTable();
renderDeploymentTable();
renderIncidentTable();
setupNavigation();
setupReleaseControls();
setupOverallStatus();
