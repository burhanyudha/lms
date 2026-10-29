(function (root) {
  function normalized(value) {
    return String(value == null ? "" : value).toLowerCase().replace(/[^a-z0-9]/g, "").trim();
  }

  function field(row, names) {
    for (const name of names) {
      const key = Object.keys(row || {}).find(candidate => normalized(candidate) === normalized(name));
      if (key && row[key] != null && row[key] !== "") return row[key];
    }
    return "";
  }

  function getDrivePreviewUrl(rawUrl) {
    try {
      const url = new URL(String(rawUrl || ""));
      if (url.protocol !== "https:") return "";

      const driveFile = url.hostname === "drive.google.com" && url.pathname.match(/^\/file\/d\/([^/]+)/);
      const workspaceFile = url.hostname === "docs.google.com" && url.pathname.match(/^\/(document|presentation|spreadsheets)\/d\/([^/]+)/);
      const fileId = driveFile ? driveFile[1] : workspaceFile ? workspaceFile[2] : url.hostname === "drive.google.com" ? url.searchParams.get("id") : "";
      if (!fileId || !/^[\w-]+$/.test(fileId)) return "";

      if (driveFile || url.hostname === "drive.google.com") return `https://drive.google.com/file/d/${fileId}/preview`;
      return `https://docs.google.com/${workspaceFile[1]}/d/${fileId}/${workspaceFile[1] === "presentation" ? "embed" : "preview"}`;
    } catch (error) {
      return "";
    }
  }

  function findStudentTaskUpload(records, student, materialId, matchesStudent) {
    const targetId = String(materialId || "").trim();
    return [...(records || [])].reverse().find(record =>
      String(field(record, ["id_materi", "idmateri", "modul", "materi"]) || "").trim() === targetId &&
      matchesStudent(record, student)
    ) || null;
  }

  function buildTaskGradePayload(task, nilai, waktu) {
    return {
      action: "updateNilaiByUsernameMateri",
      username: task.username,
      nama: task.nama,
      kelas: task.kelas,
      id_materi: task.idMateri,
      nilai: nilai,
      waktu: waktu
    };
  }

  getDrivePreviewUrl.findStudentTaskUpload = findStudentTaskUpload;
  getDrivePreviewUrl.buildTaskGradePayload = buildTaskGradePayload;
  if (typeof module === "object" && module.exports) module.exports = getDrivePreviewUrl;
  else {
    root.getDrivePreviewUrl = getDrivePreviewUrl;
    root.findStudentTaskUpload = findStudentTaskUpload;
    root.buildTaskGradePayload = buildTaskGradePayload;
  }
})(globalThis);
