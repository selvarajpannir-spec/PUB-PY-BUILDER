package com.docs.plus;

import android.app.Activity;
import android.app.AlertDialog;
import android.content.ContentValues;
import android.content.Intent;
import android.content.ContentResolver;
import android.content.pm.PackageManager;
import android.graphics.Color;
import android.net.Uri;
import android.provider.MediaStore;
import android.media.MediaScannerConnection;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.view.View;
import android.view.Window;
import android.view.WindowManager;
import android.webkit.JavascriptInterface;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.webkit.WebChromeClient;
import android.widget.Toast;
import android.Manifest;
import android.content.SharedPreferences;
import androidx.documentfile.provider.DocumentFile;
import android.provider.Settings;
import java.util.ArrayList;

import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.core.content.FileProvider;
import androidx.core.view.WindowCompat;

import com.chaquo.python.PyObject;
import com.chaquo.python.Python;
import com.chaquo.python.android.AndroidPlatform;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

public class MainActivity extends Activity {

    private WebView webView;
    private static final int PICK_EXCEL_REQUEST   = 1;
    private static final int PICK_WORD_REQUEST    = 2;
    private static final int NOTIF_PERMISSION_REQ = 3;
    private static final int STORAGE_PERMISSION_REQ = 4;

    private String cachedExcelPath = null;
    private String cachedWordPath  = null;
    private Uri analyserTreeUri = null;
    private static final int PICK_ANALYSER_TREE = 5;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        // ── Fix title overlapping status bar ──
        // 1. Let the system reserve space for status & navigation bars
        WindowCompat.setDecorFitsSystemWindows(getWindow(), true);

        // 2. Clear any fullscreen / layout-no-limits flags
        getWindow().clearFlags(WindowManager.LayoutParams.FLAG_FULLSCREEN);
        getWindow().clearFlags(WindowManager.LayoutParams.FLAG_LAYOUT_NO_LIMITS);
        getWindow().clearFlags(WindowManager.LayoutParams.FLAG_TRANSLUCENT_STATUS);
        getWindow().clearFlags(WindowManager.LayoutParams.FLAG_TRANSLUCENT_NAVIGATION);

        // 3. Solid status bar colour matching the app header
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.LOLLIPOP) {
            Window window = getWindow();
            window.addFlags(WindowManager.LayoutParams.FLAG_DRAWS_SYSTEM_BAR_BACKGROUNDS);
            window.setStatusBarColor(Color.parseColor("#1a365d"));   // same blue as header
            window.setNavigationBarColor(Color.parseColor("#f0f4f8"));
        }

        // 4. Light icons only if needed (we use dark blue, so keep light icons)
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            View decor = getWindow().getDecorView();
            // Keep default (light icons on dark status bar)
            decor.setSystemUiVisibility(0);
        }

        // Start Python
        if (!Python.isStarted()) {
            Python.start(new AndroidPlatform(this));
        }

        ensureRuntimePermissions(true);

        webView = new WebView(this);

        WebSettings s = webView.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setAllowFileAccess(true);
        s.setAllowContentAccess(true);
        s.setAllowFileAccessFromFileURLs(true);
        s.setAllowUniversalAccessFromFileURLs(true);
        s.setMixedContentMode(WebSettings.MIXED_CONTENT_ALWAYS_ALLOW);
        s.setLoadWithOverviewMode(true);
        s.setUseWideViewPort(true);

        webView.addJavascriptInterface(new AppBridge(), "AppBridge");
        webView.setWebViewClient(new WebViewClient());
        webView.setWebChromeClient(new WebChromeClient());

        webView.loadUrl("file:///android_asset/splash.html");
        setContentView(webView);
    }

    /** Open a local file path via FileProvider, or a content:// URI directly. */
    private void openDocument(String pathOrUri) {
        try {
            Uri uri;
            if (pathOrUri.startsWith("content://")) {
                uri = Uri.parse(pathOrUri);
            } else {
                File file = new File(pathOrUri);
                if (!file.exists()) {
                    Toast.makeText(this, "File not found", Toast.LENGTH_SHORT).show();
                    return;
                }
                uri = FileProvider.getUriForFile(this,
                        getPackageName() + ".provider", file);
            }
            Intent intent = new Intent(Intent.ACTION_VIEW);
            intent.setDataAndType(uri,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION
                    | Intent.FLAG_ACTIVITY_NEW_TASK);
            startActivity(Intent.createChooser(intent, "Open document with"));
        } catch (Exception e) {
            Toast.makeText(this, "Cannot open file: " + e.getMessage(),
                    Toast.LENGTH_LONG).show();
        }
    }

    private void showOpenDialog(String pathOrUri, String displayName) {
        runOnUiThread(() -> {
            String msg = "File saved to Downloads:\n"
                    + displayName
                    + "\n\nLook in your Downloads folder "
                    + "(or Files app → Downloads).\n\n"
                    + "Open it now?";
            new AlertDialog.Builder(this)
                .setTitle("Document Ready")
                .setMessage(msg)
                .setPositiveButton("Yes", (dialog, which) -> openDocument(pathOrUri))
                .setNegativeButton("No", (dialog, which) -> dialog.dismiss())
                .setCancelable(true)
                .show();
        });
    }

    /**
     * Copy a local file into the public Downloads folder so the user can find it
     * in any file manager. Returns a path or content URI string for opening.
     */
    private String publishToDownloads(File source, String fileName) throws Exception {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
            // Android 10+ — MediaStore Downloads collection
            ContentValues values = new ContentValues();
            values.put(MediaStore.Downloads.DISPLAY_NAME, fileName);
            values.put(MediaStore.Downloads.MIME_TYPE,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            values.put(MediaStore.Downloads.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS);
            values.put(MediaStore.Downloads.IS_PENDING, 1);

            ContentResolver resolver = getContentResolver();
            Uri collection = MediaStore.Downloads.EXTERNAL_CONTENT_URI;
            Uri itemUri = resolver.insert(collection, values);
            if (itemUri == null) {
                throw new Exception("Cannot create file in Downloads");
            }

            try (OutputStream os = resolver.openOutputStream(itemUri);
                 FileInputStream is = new FileInputStream(source)) {
                if (os == null) throw new Exception("Cannot write to Downloads");
                byte[] buf = new byte[8192];
                int n;
                while ((n = is.read(buf)) != -1) os.write(buf, 0, n);
                os.flush();
            }

            values.clear();
            values.put(MediaStore.Downloads.IS_PENDING, 0);
            resolver.update(itemUri, values, null, null);

            // content:// URI — open directly; also visible in Downloads
            return itemUri.toString();
        } else {
            // Android 9 and below — public Downloads directory
            File dir = Environment.getExternalStoragePublicDirectory(
                    Environment.DIRECTORY_DOWNLOADS);
            if (!dir.exists()) dir.mkdirs();
            File dest = new File(dir, fileName);
            try (FileInputStream is = new FileInputStream(source);
                 FileOutputStream os = new FileOutputStream(dest)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = is.read(buf)) != -1) os.write(buf, 0, n);
                os.flush();
            }
            MediaScannerConnection.scanFile(this,
                    new String[]{dest.getAbsolutePath()},
                    new String[]{"application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
                    null);
            return dest.getAbsolutePath();
        }
    }



    @Override
    protected void onResume() {
        super.onResume();
        // Soft re-check only — never open All-files settings (Play policy)
        ensureRuntimePermissions(false);
        restoreAnalyserTreeUri();
    }

    /**
     * Play-safe permissions only:
     * - READ_EXTERNAL_STORAGE (API ≤ 32)
     * - READ_MEDIA_IMAGES (API 33+) for photos in Excel
     * - POST_NOTIFICATIONS (API 33+)
     * No MANAGE_EXTERNAL_STORAGE.
     */
    private void ensureRuntimePermissions(boolean prompt) {
        ArrayList<String> need = new ArrayList<>();

        if (Build.VERSION.SDK_INT >= 33) {
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.POST_NOTIFICATIONS)
                    != PackageManager.PERMISSION_GRANTED) {
                need.add(Manifest.permission.POST_NOTIFICATIONS);
            }
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.READ_MEDIA_IMAGES)
                    != PackageManager.PERMISSION_GRANTED) {
                need.add(Manifest.permission.READ_MEDIA_IMAGES);
            }
        } else if (Build.VERSION.SDK_INT >= 23) {
            if (ContextCompat.checkSelfPermission(this, Manifest.permission.READ_EXTERNAL_STORAGE)
                    != PackageManager.PERMISSION_GRANTED) {
                need.add(Manifest.permission.READ_EXTERNAL_STORAGE);
            }
            if (Build.VERSION.SDK_INT <= 28
                    && ContextCompat.checkSelfPermission(this, Manifest.permission.WRITE_EXTERNAL_STORAGE)
                    != PackageManager.PERMISSION_GRANTED) {
                need.add(Manifest.permission.WRITE_EXTERNAL_STORAGE);
            }
        }

        if (prompt && !need.isEmpty()) {
            ActivityCompat.requestPermissions(this,
                    need.toArray(new String[0]), STORAGE_PERMISSION_REQ);
        }
    }

    private void restoreAnalyserTreeUri() {
        try {
            SharedPreferences sp = getSharedPreferences("docsplus", MODE_PRIVATE);
            String s = sp.getString("analyser_tree_uri", null);
            if (s != null && s.length() > 0) {
                analyserTreeUri = Uri.parse(s);
            }
        } catch (Exception ignored) {}
    }

    private void saveAnalyserTreeUri(Uri uri) {
        analyserTreeUri = uri;
        try {
            getSharedPreferences("docsplus", MODE_PRIVATE)
                    .edit()
                    .putString("analyser_tree_uri", uri != null ? uri.toString() : null)
                    .apply();
        } catch (Exception ignored) {}
    }

    private String uriToTempFile(String uriString, String suffix) {
        InputStream is = null;
        FileOutputStream fos = null;
        try {
            Uri uri = Uri.parse(uriString);
            try {
                getContentResolver().takePersistableUriPermission(uri, Intent.FLAG_GRANT_READ_URI_PERMISSION);
            } catch (Exception ignored) {}
            is = getContentResolver().openInputStream(uri);
            if (is == null) {
                return "ERROR:Cannot open file (permission denied or unsupported)";
            }
            File tmp = File.createTempFile("docgen_", suffix, getCacheDir());
            fos = new FileOutputStream(tmp);
            byte[] buf = new byte[8192];
            int n;
            long total = 0;
            while ((n = is.read(buf)) != -1) {
                fos.write(buf, 0, n);
                total += n;
            }
            fos.flush();
            if (total <= 0) {
                try { tmp.delete(); } catch (Exception ignored) {}
                return "ERROR:File empty or unreadable";
            }
            return tmp.getAbsolutePath();
        } catch (Exception e) {
            return "ERROR:" + (e.getMessage() != null ? e.getMessage() : "read failed");
        } finally {
            try { if (is != null) is.close(); } catch (Exception ignored) {}
            try { if (fos != null) fos.close(); } catch (Exception ignored) {}
        }
    }

    private void notifyJsFileResult(String jsFn, boolean ok, String msg) {
        String safe = (msg == null ? "" : msg)
                .replace("\\", "\\\\")
                .replace("'", "\\'")
                .replace("\n", " ")
                .replace("\r", "");
        final String script = jsFn + "(" + (ok ? "true" : "false") + ",'" + safe + "')";
        runOnUiThread(() -> {
            if (webView != null) webView.evaluateJavascript(script, null);
            if (!ok) {
                Toast.makeText(MainActivity.this, "File error: " + safe, Toast.LENGTH_LONG).show();
            } else {
                Toast.makeText(MainActivity.this, "File loaded", Toast.LENGTH_SHORT).show();
            }
        });
    }

    public class AppBridge {

        @JavascriptInterface
        public void pickExcelFile() {
            runOnUiThread(() -> {
                try {
                    Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
                    intent.addCategory(Intent.CATEGORY_OPENABLE);
                    intent.setType("*/*");
                    intent.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "application/vnd.ms-excel",
                        "application/octet-stream"
                    });
                    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION
                            | Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION);
                    startActivityForResult(
                            Intent.createChooser(intent, "Select Excel File"),
                            PICK_EXCEL_REQUEST);
                } catch (Exception e) {
                    Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                    intent.addCategory(Intent.CATEGORY_OPENABLE);
                    intent.setType("*/*");
                    startActivityForResult(
                            Intent.createChooser(intent, "Select Excel File"),
                            PICK_EXCEL_REQUEST);
                }
            });
        }

        @JavascriptInterface
        public void pickWordFile() {
            runOnUiThread(() -> {
                try {
                    Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
                    intent.addCategory(Intent.CATEGORY_OPENABLE);
                    intent.setType("*/*");
                    intent.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "application/octet-stream"
                    });
                    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION
                            | Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION);
                    startActivityForResult(
                            Intent.createChooser(intent, "Select Word Template"),
                            PICK_WORD_REQUEST);
                } catch (Exception e) {
                    Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                    intent.addCategory(Intent.CATEGORY_OPENABLE);
                    intent.setType("*/*");
                    startActivityForResult(
                            Intent.createChooser(intent, "Select Word Template"),
                            PICK_WORD_REQUEST);
                }
            });
        }

        @JavascriptInterface
        public String analyseFiles() {
            if (cachedExcelPath == null || cachedWordPath == null)
                return "{\"error\":\"Files not loaded\"}";
            try {
                Python py = Python.getInstance();
                PyObject gen = py.getModule("generator");

                PyObject colResult = gen.callAttr("get_column_names", cachedExcelPath);
                PyObject colsPy    = colResult.asList().get(0);
                int rowCount       = colResult.asList().get(1).toInt();

                StringBuilder colsJson = new StringBuilder("[");
                for (int i = 0; i < colsPy.asList().size(); i++) {
                    if (i > 0) colsJson.append(",");
                    String col = colsPy.asList().get(i).toString()
                                       .replace("\\", "\\\\")
                                       .replace("\"", "\\\"");
                    colsJson.append("\"").append(col).append("\"");
                }
                colsJson.append("]");

                PyObject placeholdersPy = gen.callAttr("get_placeholders", cachedWordPath);
                StringBuilder holdersJson = new StringBuilder("[");
                for (int i = 0; i < placeholdersPy.asList().size(); i++) {
                    if (i > 0) holdersJson.append(",");
                    String p = placeholdersPy.asList().get(i).toString()
                                     .replace("\\", "\\\\")
                                     .replace("\"", "\\\"");
                    holdersJson.append("\"").append(p).append("\"");
                }
                holdersJson.append("]");

                return "{\"columns\":" + colsJson +
                       ",\"rows\":" + rowCount +
                       ",\"placeholders\":" + holdersJson + "}";
            } catch (Exception e) {
                return "{\"error\":\"" + e.getMessage()
                        .replace("\\", "\\\\")
                        .replace("\"", "'") + "\"}";
            }
        }

        @JavascriptInterface
        public String generateDocx() {
            if (cachedExcelPath == null || cachedWordPath == null)
                return "ERROR:Files not loaded";
            try {
                // 1) Generate into app cache (Python can always write here)
                String fileName = "DocsPlus_" + System.currentTimeMillis() + ".docx";
                File tmpFile = new File(getCacheDir(), fileName);

                Python py  = Python.getInstance();
                PyObject gen = py.getModule("generator");
                PyObject result = gen.callAttr("generate",
                    cachedExcelPath,
                    cachedWordPath,
                    tmpFile.getAbsolutePath());

                boolean success = result.asList().get(0).toBoolean();
                String  message = result.asList().get(1).toString();
                int     count   = result.asList().get(2).toInt();

                if (!success) {
                    return "ERROR:" + message;
                }

                // 2) Copy into public Downloads so it appears in file managers
                String published = publishToDownloads(tmpFile, fileName);

                // Clean temp
                try { tmpFile.delete(); } catch (Exception ignored) {}

                showOpenDialog(published, fileName);
                return "OK:" + fileName;
            } catch (Exception e) {
                return "ERROR:" + e.getMessage();
            }
        }

        @JavascriptInterface
        public void showToast(String msg) {
            runOnUiThread(() ->
                Toast.makeText(MainActivity.this, msg, Toast.LENGTH_SHORT).show());
        }

        @JavascriptInterface
        public void clearFiles() {
            if (cachedExcelPath != null) new File(cachedExcelPath).delete();
            if (cachedWordPath  != null) new File(cachedWordPath).delete();
            cachedExcelPath = null;
            cachedWordPath  = null;
        }

        // ── Generator batch API ──

        @JavascriptInterface
        public String listExcelRows() {
            try {
                if (cachedExcelPath == null) return "[]";
                ensurePython();
                Python py = Python.getInstance();
                PyObject gen = py.getModule("generator");
                return gen.callAttr("list_data_rows", cachedExcelPath).toString();
            } catch (Exception e) {
                return "{\"error\":\"" + safe(e.getMessage()) + "\"}";
            }
        }

        @JavascriptInterface
        public String generateOneRow(String excelRowNum) {
            try {
                if (cachedExcelPath == null || cachedWordPath == null)
                    return "ERROR:Files not loaded";
                ensurePython();
                File out = new File(getCacheDir(),
                        "row_" + excelRowNum + "_" + System.currentTimeMillis() + ".docx");
                Python py = Python.getInstance();
                PyObject gen = py.getModule("generator");
                PyObject result = gen.callAttr(
                        "generate_one_row",
                        cachedExcelPath,
                        cachedWordPath,
                        excelRowNum,
                        out.getAbsolutePath());
                boolean ok = result.asList().get(0).toBoolean();
                String msg = result.asList().get(1).toString();
                if (!ok) return "ERROR:" + msg;
                return "OK:" + out.getAbsolutePath();
            } catch (Exception e) {
                return "ERROR:" + e.getMessage();
            }
        }

        @JavascriptInterface
        public String mergeAndPublishBatch(String pathsJson) {
            try {
                ensurePython();
                String fileName = "DocsPlus_Batch_" + System.currentTimeMillis() + ".docx";
                File merged = new File(getCacheDir(), fileName);
                Python py = Python.getInstance();
                PyObject gen = py.getModule("generator");
                PyObject result = gen.callAttr(
                        "merge_docx_list",
                        pathsJson,
                        merged.getAbsolutePath());
                boolean ok = result.asList().get(0).toBoolean();
                String msg = result.asList().get(1).toString();
                if (!ok) return "ERROR:" + msg;

                String published = publishToDownloads(merged, fileName);
                try { merged.delete(); } catch (Exception ignored) {}
                // cleanup row temps
                try {
                    String s = pathsJson.replace("[", "").replace("]", "").replace("\"", "");
                    for (String p : s.split(",")) {
                        p = p.trim();
                        if (p.length() > 0) {
                            File f = new File(p);
                            if (f.exists()) f.delete();
                        }
                    }
                } catch (Exception ignored) {}
                return "OK:" + fileName + "|" + published;
            } catch (Exception e) {
                return "ERROR:" + e.getMessage();
            }
        }

        // ── Document Analyser API ──


        /** User must pick a folder via SAF (Play-safe). */
        @JavascriptInterface
        public void pickAnalyserFolder() {
            runOnUiThread(() -> {
                Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT_TREE);
                intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION
                        | Intent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION
                        | Intent.FLAG_GRANT_PREFIX_URI_PERMISSION);
                startActivityForResult(intent, PICK_ANALYSER_TREE);
            });
        }

        @JavascriptInterface
        public String getAnalyserFolder() {
            restoreAnalyserTreeUri();
            if (analyserTreeUri == null) {
                return "(no folder selected — tap Choose folder)";
            }
            try {
                DocumentFile tree = DocumentFile.fromTreeUri(MainActivity.this, analyserTreeUri);
                if (tree != null && tree.getName() != null) {
                    return tree.getName();
                }
            } catch (Exception ignored) {}
            return analyserTreeUri.toString();
        }

        /**
         * List documents under the user-picked SAF tree.
         * Copies each file into cache so Python can read a real path.
         */
        @JavascriptInterface
        public String listAnalyserFiles() {
            restoreAnalyserTreeUri();
            if (analyserTreeUri == null) {
                return "{\"error\":\"Choose a folder first (SAF). All-files access is not used.\"}";
            }
            try {
                DocumentFile tree = DocumentFile.fromTreeUri(MainActivity.this, analyserTreeUri);
                if (tree == null || !tree.isDirectory()) {
                    return "{\"error\":\"Folder unavailable. Choose folder again.\"}";
                }
                String[] exts = { ".pdf", ".txt", ".md", ".xlsx", ".docx" };
                StringBuilder sb = new StringBuilder("{\"files\":[");
                int n = 0;
                n = MainActivity.this.appendDocumentFiles(tree, exts, sb, n, 0);
                sb.append("]}");
                return sb.toString();
            } catch (Exception e) {
                return "{\"error\":" + jsonStr(e.getMessage()) + "}";
            }
        }

        @JavascriptInterface
        public String analyseOneFile(String path) {
            try {
                if (path == null || path.length() == 0) {
                    return "{\"ok\":false,\"name\":\"\",\"path\":\"\",\"preview\":\"\",\"tags\":[],\"note\":\"empty path\",\"error\":\"empty path\"}";
                }
                ensurePython();
                Python py = Python.getInstance();
                PyObject mod = py.getModule("analyser");
                return mod.callAttr("analyse_one", path).toString();
            } catch (Exception e) {
                return "{\"ok\":false,\"name\":" + jsonStr(new File(path).getName())
                        + ",\"path\":" + jsonStr(path)
                        + ",\"preview\":\"\",\"tags\":[],\"note\":"
                        + jsonStr(e.getMessage())
                        + ",\"error\":" + jsonStr(e.getMessage()) + "}";
            }
        }

        @JavascriptInterface
        public void openAnalysedFile(String path) {
            runOnUiThread(() -> {
                try {
                    if (path != null && path.startsWith("content://")) {
                        Uri uri = Uri.parse(path);
                        Intent intent = new Intent(Intent.ACTION_VIEW);
                        intent.setDataAndType(uri, mimeForName(path));
                        intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION
                                | Intent.FLAG_ACTIVITY_NEW_TASK);
                        startActivity(Intent.createChooser(intent, "Open with"));
                        return;
                    }
                    File file = new File(path);
                    if (!file.exists()) {
                        Toast.makeText(MainActivity.this, "File not found", Toast.LENGTH_SHORT).show();
                        return;
                    }
                    Uri uri = FileProvider.getUriForFile(
                            MainActivity.this,
                            getPackageName() + ".provider",
                            file);
                    Intent intent = new Intent(Intent.ACTION_VIEW);
                    intent.setDataAndType(uri, mimeForName(file.getName()));
                    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION
                            | Intent.FLAG_ACTIVITY_NEW_TASK);
                    startActivity(Intent.createChooser(intent, "Open with"));
                } catch (Exception e) {
                    Toast.makeText(MainActivity.this,
                            "Cannot open: " + e.getMessage(), Toast.LENGTH_LONG).show();
                }
            });
        }

        @JavascriptInterface
        public String deleteStorageFile(String path) {
            try {
                File f = new File(path);
                if (!f.exists()) return "OK";
                if (f.delete()) return "OK";
                return "ERROR:delete failed";
            } catch (Exception e) {
                return "ERROR:" + e.getMessage();
            }
        }
    }


    private void ensurePython() {
        if (!Python.isStarted()) {
            Python.start(new AndroidPlatform(this));
        }
    }

    private static String safe(String s) {
        if (s == null) return "";
        return s.replace("\\", "\\\\").replace("\"", "'").replace("\n", " ");
    }


    private String getAnalyserFolderLabel(Uri uri) {
        try {
            DocumentFile tree = DocumentFile.fromTreeUri(this, uri);
            if (tree != null && tree.getName() != null) {
                return tree.getName();
            }
        } catch (Exception ignored) {}
        return uri != null ? uri.toString() : "";
    }

    private int appendDocumentFiles(DocumentFile dir, String[] exts, StringBuilder sb, int n, int depth) {
        if (dir == null || n >= 200 || depth > 2) return n;
        DocumentFile[] kids = dir.listFiles();
        if (kids == null) return n;
        for (DocumentFile f : kids) {
            if (n >= 200) break;
            if (f.isDirectory()) {
                n = appendDocumentFiles(f, exts, sb, n, depth + 1);
                continue;
            }
            String name = f.getName();
            if (name == null || !matchExt(name, exts)) continue;
            String path = copyDocumentToCache(f, name);
            if (path == null || path.startsWith("ERROR:")) continue;
            if (n > 0) sb.append(',');
            sb.append("{\"name\":").append(jsonStr(name))
              .append(",\"path\":").append(jsonStr(path))
              .append('}');
            n++;
        }
        return n;
    }

    private String copyDocumentToCache(DocumentFile doc, String name) {
        InputStream is = null;
        FileOutputStream fos = null;
        try {
            Uri uri = doc.getUri();
            is = getContentResolver().openInputStream(uri);
            if (is == null) return "ERROR:stream";
            String suffix = "";
            int dot = name.lastIndexOf('.');
            if (dot >= 0) suffix = name.substring(dot);
            File tmp = File.createTempFile("an_", suffix, getCacheDir());
            fos = new FileOutputStream(tmp);
            byte[] buf = new byte[8192];
            int r;
            long total = 0;
            while ((r = is.read(buf)) != -1) {
                fos.write(buf, 0, r);
                total += r;
            }
            fos.flush();
            if (total <= 0) {
                try { tmp.delete(); } catch (Exception ignored) {}
                return "ERROR:empty";
            }
            return tmp.getAbsolutePath();
        } catch (Exception e) {
            return "ERROR:" + (e.getMessage() != null ? e.getMessage() : "copy failed");
        } finally {
            try { if (is != null) is.close(); } catch (Exception ignored) {}
            try { if (fos != null) fos.close(); } catch (Exception ignored) {}
        }
    }

    private static boolean matchExt(String name, String[] exts) {
        String low = name.toLowerCase();
        for (String e : exts) {
            if (low.endsWith(e)) return true;
        }
        return false;
    }

    private static String jsonStr(String s) {
        if (s == null) s = "";
        return "\"" + s.replace("\\", "\\\\")
                       .replace("\"", "\\\"")
                       .replace("\n", "\\n")
                       .replace("\r", "") + "\"";
    }

    private static String mimeForName(String name) {
        String low = name.toLowerCase();
        if (low.endsWith(".pdf")) return "application/pdf";
        if (low.endsWith(".txt")) return "text/plain";
        if (low.endsWith(".md"))  return "text/markdown";
        if (low.endsWith(".xlsx"))
            return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        if (low.endsWith(".docx"))
            return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
        return "*/*";
    }

    @Override
    protected void onActivityResult(int req, int res, Intent data) {
        super.onActivityResult(req, res, data);
        if (res != RESULT_OK) {
            Toast.makeText(this, "Selection cancelled", Toast.LENGTH_SHORT).show();
            return;
        }
        if (data == null || data.getData() == null) {
            Toast.makeText(this, "No file/folder returned", Toast.LENGTH_LONG).show();
            return;
        }

        final Uri uri = data.getData();

        if (req == PICK_ANALYSER_TREE) {
            try {
                final int takeFlags = data.getFlags()
                        & (Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
                getContentResolver().takePersistableUriPermission(uri,
                        takeFlags != 0 ? takeFlags : Intent.FLAG_GRANT_READ_URI_PERMISSION);
            } catch (Exception ignored) {}
            saveAnalyserTreeUri(uri);
            String label = getAnalyserFolderLabel(uri);
            Toast.makeText(this, "Folder selected: " + label, Toast.LENGTH_LONG).show();
            final String safe = (label == null ? "" : label)
                    .replace("\\", "\\\\")
                    .replace("'", "\\'");
            if (webView != null) {
                webView.evaluateJavascript(
                        "if(window.onAnalyserFolderReady)onAnalyserFolderReady('" + safe + "')",
                        null);
            }
            return;
        }
        try {
            int flags = data.getFlags()
                    & (Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
            if (flags == 0) flags = Intent.FLAG_GRANT_READ_URI_PERMISSION;
            getContentResolver().takePersistableUriPermission(uri, flags);
        } catch (Exception ignored) {}

        final String uriStr = uri.toString();

        if (req == PICK_EXCEL_REQUEST) {
            Toast.makeText(this, "Loading Excel…", Toast.LENGTH_SHORT).show();
            new Thread(() -> {
                String path = uriToTempFile(uriStr, ".xlsx");
                if (path != null && path.startsWith("ERROR:")) {
                    cachedExcelPath = null;
                    notifyJsFileResult("onExcelReady", false, path);
                } else {
                    cachedExcelPath = path;
                    notifyJsFileResult("onExcelReady", true, uriStr);
                }
            }).start();

        } else if (req == PICK_WORD_REQUEST) {
            Toast.makeText(this, "Loading Word template…", Toast.LENGTH_SHORT).show();
            new Thread(() -> {
                String path = uriToTempFile(uriStr, ".docx");
                if (path != null && path.startsWith("ERROR:")) {
                    cachedWordPath = null;
                    notifyJsFileResult("onWordReady", false, path);
                } else {
                    cachedWordPath = path;
                    notifyJsFileResult("onWordReady", true, uriStr);
                }
            }).start();
        }
    }

    @Override
    public void onRequestPermissionsResult(int requestCode, String[] permissions, int[] grantResults) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults);
        if (requestCode == STORAGE_PERMISSION_REQ) {
            boolean granted = grantResults != null && grantResults.length > 0;
            if (granted) {
                for (int g : grantResults) {
                    if (g != PackageManager.PERMISSION_GRANTED) {
                        granted = false;
                        break;
                    }
                }
            }
            if (!granted) {
                Toast.makeText(this,
                        "Storage permission denied — file pick may fail",
                        Toast.LENGTH_LONG).show();
            }
        }
    }

    @Override
    public void onBackPressed() {
        if (webView != null && webView.canGoBack()) {
            webView.goBack();
            return;
        }

        new AlertDialog.Builder(this)
            .setTitle("Exit App")
            .setMessage("Do you want to exit Docs+?")
            .setPositiveButton("Exit", (dialog, which) -> finishAffinity())
            .setNegativeButton("Cancel", (dialog, which) -> dialog.dismiss())
            .setCancelable(true)
            .show();
    }
}
