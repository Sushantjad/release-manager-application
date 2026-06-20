const http = require("node:http");
const fs = require("node:fs/promises");
const path = require("node:path");

const HOST = "127.0.0.1";
const PORT = Number(process.env.PORT) || 3000;
const ROOT = __dirname;
const DATA_FILE = path.join(ROOT, "data", "release-data.json");

const contentTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png"
};

function sendJson(response, statusCode, body) {
  response.writeHead(statusCode, { "Content-Type": contentTypes[".json"] });
  response.end(JSON.stringify(body, null, 2));
}

async function readReleaseData() {
  return JSON.parse(await fs.readFile(DATA_FILE, "utf8"));
}

async function writeReleaseData(data) {
  await fs.writeFile(DATA_FILE, `${JSON.stringify(data, null, 2)}\n`, "utf8");
}

async function readRequestBody(request) {
  let body = "";
  for await (const chunk of request) {
    body += chunk;
    if (body.length > 1_000_000) {
      throw new Error("Request body is too large");
    }
  }
  return body ? JSON.parse(body) : {};
}

async function handleApi(request, response, pathname) {
  if (request.method === "GET" && pathname === "/api/health") {
    return sendJson(response, 200, {
      status: "ok",
      service: "release-manager-api",
      time: new Date().toISOString()
    });
  }

  if (request.method === "GET" && pathname === "/api/releases/current") {
    return sendJson(response, 200, await readReleaseData());
  }

  if (request.method === "PATCH" && pathname === "/api/releases/current/status") {
    const update = await readRequestBody(request);
    const validStates = ["Not Started", "Release Started", "Release Stopped"];
    const validHealth = ["Green", "Yellow", "Red"];

    if (update.state && !validStates.includes(update.state)) {
      return sendJson(response, 400, { error: "Invalid release state" });
    }
    if (update.overallStatus && !validHealth.includes(update.overallStatus)) {
      return sendJson(response, 400, { error: "Invalid overall status" });
    }

    const data = await readReleaseData();
    if (update.state) data.release.state = update.state;
    if (update.overallStatus) data.release.overallStatus = update.overallStatus;
    data.release.updatedAt = new Date().toISOString();
    await writeReleaseData(data);
    return sendJson(response, 200, data.release);
  }

  return sendJson(response, 404, { error: "API route not found" });
}

async function serveStaticFile(response, pathname) {
  const requestedPath = pathname === "/" ? "/index.html" : pathname;
  const filePath = path.resolve(ROOT, `.${requestedPath}`);

  if (!filePath.startsWith(`${ROOT}${path.sep}`)) {
    return sendJson(response, 403, { error: "Forbidden" });
  }

  try {
    const file = await fs.readFile(filePath);
    const contentType = contentTypes[path.extname(filePath)] || "application/octet-stream";
    response.writeHead(200, { "Content-Type": contentType });
    response.end(file);
  } catch (error) {
    if (error.code === "ENOENT") {
      return sendJson(response, 404, { error: "File not found" });
    }
    throw error;
  }
}

const server = http.createServer(async (request, response) => {
  try {
    const url = new URL(request.url, `http://${request.headers.host}`);
    if (url.pathname.startsWith("/api/")) {
      await handleApi(request, response, url.pathname);
    } else {
      await serveStaticFile(response, decodeURIComponent(url.pathname));
    }
  } catch (error) {
    console.error(error);
    sendJson(response, 500, { error: "Internal server error" });
  }
});

server.listen(PORT, HOST, () => {
  console.log(`Release Manager running at http://${HOST}:${PORT}`);
  console.log(`API health check: http://${HOST}:${PORT}/api/health`);
});
