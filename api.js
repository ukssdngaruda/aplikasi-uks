// Pengganti google.script.run untuk GitHub Pages.
// Kode lama seperti google.script.run.withSuccessHandler(f).withFailureHandler(g).getSheetData(token, 'Siswa')
// tetap jalan tanpa diubah — di balik layar dikirim lewat fetch() ke Apps Script.
(function () {
  function callApi(fn, args) {
    var url = (window.UKS_CONFIG || {}).API_URL;
    if (!url || url.indexOf('GANTI_DENGAN') > -1) {
      return Promise.reject(new Error('API_URL belum diisi di config.js'));
    }
    return fetch(url, {
      method: 'POST',
      // text/plain agar browser tidak mengirim preflight CORS (Apps Script tidak mendukung OPTIONS)
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ fn: fn, args: args }),
      redirect: 'follow'
    }).then(function (res) {
      if (!res.ok) throw new Error('Server error (' + res.status + ')');
      return res.json();
    }).then(function (json) {
      if (!json || json.ok !== true) throw new Error((json && json.error) || 'Respons server tidak valid');
      return json.data;
    }).catch(function (err) {
      if (err instanceof TypeError) throw new Error('Tidak dapat terhubung ke server. Periksa koneksi internet / API_URL.');
      throw err;
    });
  }

  function makeRunner(ok, fail) {
    return new Proxy({}, {
      get: function (_, prop) {
        if (prop === 'withSuccessHandler') return function (f) { return makeRunner(f, fail); };
        if (prop === 'withFailureHandler') return function (f) { return makeRunner(ok, f); };
        if (prop === 'withUserObject') return function () { return makeRunner(ok, fail); };
        return function () {
          var args = Array.prototype.slice.call(arguments);
          callApi(String(prop), args).then(
            function (data) { if (ok) ok(data); },
            function (err) { if (fail) fail(err); else console.error(err); }
          );
        };
      }
    });
  }

  window.google = window.google || {};
  window.google.script = { run: makeRunner(null, null) };
})();
