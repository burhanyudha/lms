(function () {
  const API_PREFIX = "https://script.google.com/macros/s/";
  const nativeFetch = window.fetch.bind(window);

  function endpointFrom(url) {
    return url.split("?")[0];
  }

  function sessionToken() {
    const keys = /admin\.html$/.test(location.pathname) || (/login-admin\.html$/.test(location.pathname) && localStorage.getItem("admin_user")) ? ["admin_user"] : ["user"];
    for (const key of keys) {
      try {
        const session = JSON.parse(localStorage.getItem(key) || "null");
        if (session && session.token) return session.token;
      } catch (_) {}
    }
    return "";
  }

  function clearSession() {
    const key = /admin\.html$/.test(location.pathname) ? "admin_user" : "user";
    [key, "cache_user", "cache_materi", "cache_listNilai", "cache_logTugas", "cache_guru", "cache_lastUpdated"].forEach(item => localStorage.removeItem(item));
    if (!/index\.html$|login-admin\.html$/.test(location.pathname)) location.href = "index.html";
  }

  function requireSession() {
    const key = /admin\.html$/.test(location.pathname) ? "admin_user" : "user";
    if (!localStorage.getItem(key) || !sessionToken()) location.href = /admin\.html$/.test(location.pathname) ? "login-admin.html" : "index.html";
  }

  async function checkAuthentication(response) {
    const body = await response.clone().json().catch(() => null);
    if (body && body.code === "UNAUTHORIZED") clearSession();
    if (body && body.status === "error") throw new Error(body.message || "Permintaan gagal.");
    return response;
  }

  window.lmsApi = {
    async call(action, data) {
      return nativeFetch(API_URL, {
        method: "POST",
        headers: {"Content-Type": "text/plain;charset=utf-8"},
        body: JSON.stringify({action, ...data, token: sessionToken()})
      }).then(checkAuthentication);
    },
    async login(action, identity, password) {
      const response = await nativeFetch(API_URL, {
        method: "POST",
        headers: {"Content-Type": "text/plain;charset=utf-8"},
        body: JSON.stringify({action, ...(typeof identity === "string" ? {username: identity} : identity), password})
      });
      return response.json();
    },
    async validateSession() {
      const response = await this.call("sessionInfo", {});
      return response.json();
    }
  };

  if (!/index\.html$|login-admin\.html$/.test(location.pathname)) requireSession();

  window.fetch = async function (input, init) {
    const url = typeof input === "string" ? input : input.url;
    if (!url.startsWith(API_PREFIX)) return nativeFetch(input, init);
    const options = {...(init || {})};
    const method = (options.method || (input instanceof Request ? input.method : "GET")).toUpperCase();
    const token = sessionToken();
    if (method === "GET") {
      const parsed = new URL(url);
      const sheet = parsed.searchParams.get("sheet");
      if (!sheet) return nativeFetch(input, init);
      options.method = "POST";
      options.headers = {"Content-Type": "text/plain;charset=utf-8", ...(options.headers || {})};
      options.body = JSON.stringify({action: "getData", sheet, token});
      const response = await nativeFetch(endpointFrom(url), options);
      return checkAuthentication(response);
    }
    if (typeof options.body === "string") {
      try {
        const payload = JSON.parse(options.body);
        if (payload.action !== "login" && payload.action !== "loginAdmin") payload.token = token;
        options.body = JSON.stringify(payload);
      } catch (_) {}
    }
    return checkAuthentication(await nativeFetch(input, options));
  };
})();
