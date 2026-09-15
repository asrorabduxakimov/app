package uz.attestation.excel;

import android.content.Context;
import android.graphics.Bitmap;
import androidx.test.core.app.ActivityScenario;
import androidx.test.ext.junit.runners.AndroidJUnit4;
import androidx.test.platform.app.InstrumentationRegistry;
import androidx.core.content.FileProvider;
import com.chaquo.python.Python;
import com.chaquo.python.android.AndroidPlatform;
import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.Test;
import org.junit.runner.RunWith;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipFile;
import static org.junit.Assert.*;

@RunWith(AndroidJUnit4.class)
public class AppSmokeTest {
    @Test public void pythonProcessesExcelArchivesAndSharesOnAndroid() throws Exception {
        Context context = InstrumentationRegistry.getInstrumentation().getTargetContext();
        if (!Python.isStarted()) Python.start(new AndroidPlatform(context));
        File csv = new File(context.getFilesDir(), "smoke.csv");
        try (OutputStream out = new FileOutputStream(csv)) {
            out.write(("PINFL,FIO,Savol ID,Yuklangan variant,Natija\n"
                + "30000000000001,Test A,1,A,1\n30000000000001,Test A,2,B,0\n"
                + "40000000000002,Test B,1,A,1\n40000000000002,Test B,2,C,1\n").getBytes(StandardCharsets.UTF_8));
        }
        JSONArray files = new JSONArray().put(new JSONObject().put("name", "test.csv").put("path", csv.getAbsolutePath()));
        String root = new File(context.getFilesDir(), "results").getAbsolutePath();
        JSONObject report = new JSONObject(Python.getInstance().getModule("mobile_api")
            .callAttr("process_batch", files.toString(), "{\"add_scores\":true}", root).toString());
        assertEquals(1, report.getInt("success_count"));
        assertEquals(2, report.getJSONArray("results").getJSONObject(0).getInt("participants"));
        File archive = new File(report.getString("archive"));
        try (ZipFile zip = new ZipFile(archive)) { assertNotNull(zip.getEntry("test_processed.xlsx")); }
        android.net.Uri uri = FileProvider.getUriForFile(context, context.getPackageName() + ".files", archive);
        try (InputStream in = context.getContentResolver().openInputStream(uri)) { assertEquals('P', in.read()); }

        try (ActivityScenario<MainActivity> scenario = ActivityScenario.launch(MainActivity.class)) {
            scenario.onActivity(activity -> assertNotNull(activity.findViewById(android.R.id.content)));
            scenario.recreate();
            InstrumentationRegistry.getInstrumentation().waitForIdleSync();
            Bitmap screenshot = InstrumentationRegistry.getInstrumentation().getUiAutomation().takeScreenshot();
            assertNotNull(screenshot);
            try (OutputStream out = new FileOutputStream(new File(context.getExternalFilesDir(null), "preview.png"))) {
                screenshot.compress(Bitmap.CompressFormat.PNG, 100, out);
            }
        }
    }
}
