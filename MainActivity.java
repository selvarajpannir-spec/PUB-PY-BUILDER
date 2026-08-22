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

    private String cachedExcelPath = null;
    private String cachedWordPath  = null;

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

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.TIRAMISU) {
            if (ContextCompat.checkSelfPermission(this,
                    Manifest.permission.POST_NOTIFICATIONS)
                    != PackageManager.PERMISSION_GRANTED) {
                ActivityCompat.requestPermissions(this,
                        new String[]{Manifest.permission.POST_NOTIFICATIONS},
                        NOTIF_PERMISSION_REQ);
            }
        }

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

    private String uriToTempFile(String uriString, String suffix) {
        try {
            Uri uri = Uri.parse(uriString);
            InputStream is = getContentResolver().openInputStream(uri);
            File tmp = File.createTempFile("docgen_", suffix, getCacheDir());
            FileOutputStream fos = new FileOutputStream(tmp);
            byte[] buf = new byte[8192]; int n;
            while ((n = is.read(buf)) != -1) fos.write(buf, 0, n);
            is.close(); fos.close();
            return tmp.getAbsolutePath();
        } catch (Exception e) { return "ERROR:" + e.getMessage(); }
    }

    public class AppBridge {

        @JavascriptInterface
        public void pickExcelFile() {
            Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
            intent.setType("*/*");
            intent.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel"
            });
            startActivityForResult(
                Intent.createChooser(intent, "Select Excel File"),
                PICK_EXCEL_REQUEST);
        }

        @JavascriptInterface
        public void pickWordFile() {
            Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
            intent.setType("*/*");
            intent.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            });
            startActivityForResult(
                Intent.createChooser(intent, "Select Word Template"),
                PICK_WORD_REQUEST);
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
    }

    @Override
    protected void onActivityResult(int req, int res, Intent data) {
        super.onActivityResult(req, res, data);
        if (res != RESULT_OK || data == null || data.getData() == null) return;
        String uri = data.getData().toString();

        if (req == PICK_EXCEL_REQUEST) {
            new Thread(() -> {
                String path = uriToTempFile(uri, ".xlsx");
                cachedExcelPath = path.startsWith("ERROR:") ? null : path;
                webView.post(() -> webView.evaluateJavascript(
                    "onExcelReady(" + (cachedExcelPath != null ? "true" : "false") +
                    ",'" + uri.replace("'","\\'") + "')", null));
            }).start();

        } else if (req == PICK_WORD_REQUEST) {
            new Thread(() -> {
                String path = uriToTempFile(uri, ".docx");
                cachedWordPath = path.startsWith("ERROR:") ? null : path;
                webView.post(() -> webView.evaluateJavascript(
                    "onWordReady(" + (cachedWordPath != null ? "true" : "false") +
                    ",'" + uri.replace("'","\\'") + "')", null));
            }).start();
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
