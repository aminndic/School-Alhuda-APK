package com.aminndic.schoolalhuda;

import android.app.Activity;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;

import com.chaquo.python.PyObject;
import com.chaquo.python.Python;
import com.chaquo.python.android.AndroidPlatform;

public class MainActivity extends Activity {
    private WebView webView;
    private final Handler handler = new Handler(Looper.getMainLooper());
    private boolean serverReady = false;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        webView = new WebView(this);
        setContentView(webView);

        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setAllowFileAccess(true);
        settings.setAllowContentAccess(true);
        settings.setDatabaseEnabled(true);
        settings.setBuiltInZoomControls(false);
        settings.setDisplayZoomControls(false);
        settings.setLoadWithOverviewMode(true);
        settings.setUseWideViewPort(true);

        webView.setWebChromeClient(new WebChromeClient());
        webView.setWebViewClient(new WebViewClient() {
            @Override
            public void onReceivedError(WebView view, WebResourceRequest request, WebResourceError error) {
                if (request.isForMainFrame()) {
                    waitForServer();
                }
            }
        });

        startPythonServer();
    }

    private void startPythonServer() {
        try {
            if (!Python.isStarted()) {
                Python.start(new AndroidPlatform(this));
            }
            Python py = Python.getInstance();
            PyObject module = py.getModule("android_server");
            module.callAttr("start_server", getFilesDir().getAbsolutePath());
            waitForServer();
        } catch (Exception e) {
            Toast.makeText(this, "تعذر تشغيل النظام: " + e.getMessage(), Toast.LENGTH_LONG).show();
        }
    }

    private void waitForServer() {
        if (serverReady) return;
        handler.postDelayed(() -> {
            serverReady = true;
            webView.loadUrl("http://127.0.0.1:5000/");
        }, 1500);
    }

    @Override
    public void onBackPressed() {
        if (webView.canGoBack()) {
            webView.goBack();
        } else {
            super.onBackPressed();
        }
    }

    @Override
    protected void onDestroy() {
        handler.removeCallbacksAndMessages(null);
        if (webView != null) {
            webView.destroy();
        }
        super.onDestroy();
    }
}
