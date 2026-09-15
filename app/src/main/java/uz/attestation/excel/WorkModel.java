package uz.attestation.excel;

import android.app.Application;
import android.content.ContentResolver;
import android.database.Cursor;
import android.net.Uri;
import android.os.Handler;
import android.os.Looper;
import android.provider.OpenableColumns;
import androidx.lifecycle.AndroidViewModel;
import androidx.lifecycle.MutableLiveData;
import com.chaquo.python.PyObject;
import com.chaquo.python.Python;
import com.chaquo.python.android.AndroidPlatform;
import org.json.JSONArray;
import org.json.JSONObject;
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class WorkModel extends AndroidViewModel {
    public final MutableLiveData<Integer> changed = new MutableLiveData<>(0);
    public final ArrayList<JSONObject> inputs = new ArrayList<>();
    public JSONObject report = new JSONObject();
    public JSONObject inspection = new JSONObject();
    public String status = "Добавьте файлы для обработки";
    public boolean busy = false;
    public int progress = 0;
    private final Handler main = new Handler(Looper.getMainLooper());
    private final ExecutorService worker = Executors.newSingleThreadExecutor();

    public WorkModel(Application application) {
        super(application);
        try {
            String saved = application.getSharedPreferences("app", 0).getString("inputs", "[]");
            JSONArray array = new JSONArray(saved);
            for (int i = 0; i < array.length(); i++) {
                JSONObject item = array.getJSONObject(i);
                if (new File(item.getString("path")).isFile()) inputs.add(item);
            }
            File latest = new File(application.getFilesDir(), "results/latest.json");
            if (latest.isFile()) {
                try (InputStream in = new FileInputStream(latest)) {
                    report = new JSONObject(readText(in));
                }
            }
        } catch (Exception e) {
            status = "Не удалось восстановить предыдущий сеанс: " + e.getMessage();
        }
    }

    private void notifyUi() { changed.setValue(changed.getValue() + 1); }
    public void saveInputs() {
        getApplication().getSharedPreferences("app", 0).edit()
            .putString("inputs", new JSONArray(inputs).toString()).apply();
        notifyUi();
    }

    private PyObject api() {
        if (!Python.isStarted()) Python.start(new AndroidPlatform(getApplication()));
        return Python.getInstance().getModule("mobile_api");
    }

    private void begin(String text) { busy = true; progress = 0; status = text; notifyUi(); }
    private void fail(Exception e) {
        main.post(() -> { busy = false; status = "Ошибка: " + e.getMessage(); notifyUi(); });
    }

    public void importFiles(List<Uri> uris) {
        if (busy) return;
        begin("Копирование выбранных файлов…");
        worker.execute(() -> {
            ArrayList<JSONObject> added = new ArrayList<>();
            ArrayList<String> errors = new ArrayList<>();
            ContentResolver resolver = getApplication().getContentResolver();
            for (Uri uri : uris) {
                File target = null;
                try {
                    String name = "file.xlsx";
                    try (Cursor c = resolver.query(uri, new String[]{OpenableColumns.DISPLAY_NAME}, null, null, null)) {
                        if (c != null && c.moveToFirst() && !c.isNull(0)) name = c.getString(0);
                    }
                    String lower = name.toLowerCase(Locale.ROOT);
                    if (!(lower.endsWith(".xlsx") || lower.endsWith(".xlsm") || lower.endsWith(".xls") || lower.endsWith(".csv"))) {
                        errors.add(name + ": неподдерживаемый формат");
                        continue;
                    }
                    String extension = lower.substring(lower.lastIndexOf('.'));
                    File dir = new File(getApplication().getFilesDir(), "imports");
                    if (!dir.isDirectory() && !dir.mkdirs()) throw new IOException("Не удалось создать папку импорта");
                    target = new File(dir, UUID.randomUUID() + extension);
                    try (InputStream in = resolver.openInputStream(uri); OutputStream out = new FileOutputStream(target)) {
                        if (in == null) throw new IOException("Файл недоступен");
                        copy(in, out);
                    }
                    added.add(new JSONObject().put("name", name).put("path", target.getAbsolutePath()));
                } catch (Exception e) {
                    if (target != null) target.delete();
                    errors.add(String.valueOf(e.getMessage()));
                }
            }
            main.post(() -> {
                inputs.addAll(added); busy = false;
                status = "Добавлено файлов: " + added.size() + (errors.isEmpty() ? "" : "\n" + String.join("\n", errors));
                saveInputs();
            });
        });
    }

    public void inspect(String sortColumn) {
        if (busy || inputs.isEmpty()) return;
        String path = inputs.get(0).optString("path");
        begin("Чтение столбцов и групп первого файла…");
        worker.execute(() -> {
            try {
                JSONObject result = new JSONObject(api().callAttr("inspect_file", path, sortColumn).toString());
                main.post(() -> { inspection = result; busy = false; status = "Столбцы загружены. Строк: " + result.optInt("rows"); notifyUi(); });
            } catch (Exception e) { fail(e); }
        });
    }

    public void process(JSONObject options) {
        if (busy || inputs.isEmpty()) return;
        String items = new JSONArray(inputs).toString();
        String settings = options.toString();
        begin("Подготовка обработки…");
        worker.execute(() -> {
            try {
                String root = new File(getApplication().getFilesDir(), "results").getAbsolutePath();
                JSONObject result = new JSONObject(api().callAttr("process_batch", items, settings, root, this).toString());
                main.post(() -> {
                    report = result; busy = false; progress = 100;
                    status = "Готово: " + result.optInt("success_count") + " из " + result.optInt("total") + " файлов";
                    notifyUi();
                });
            } catch (Exception e) { fail(e); }
        });
    }

    // Public because Chaquopy calls this method from Python on the worker thread.
    public void onProgress(int done, int total, String filename) {
        main.post(() -> { progress = total == 0 ? 0 : done * 100 / total;
            status = "Обработка " + done + "/" + total + " · " + filename; notifyUi(); });
    }

    public void exportZip(Uri destination) {
        if (busy) return;
        String path = report.optString("archive");
        begin("Сохранение ZIP…");
        worker.execute(() -> {
            try (InputStream in = new FileInputStream(path);
                 OutputStream out = getApplication().getContentResolver().openOutputStream(destination, "wt")) {
                if (out == null) throw new IOException("Не удалось открыть место сохранения");
                copy(in, out);
                main.post(() -> { busy = false; status = "ZIP сохранён в выбранное место"; notifyUi(); });
            } catch (Exception e) { fail(e); }
        });
    }

    public void exportFiles(Uri destination) {
        if (busy) return;
        JSONArray results = report.optJSONArray("results");
        begin("Сохранение Excel-файлов…");
        worker.execute(() -> {
            try {
                androidx.documentfile.provider.DocumentFile folder =
                    androidx.documentfile.provider.DocumentFile.fromTreeUri(getApplication(), destination);
                if (folder == null) throw new IOException("Папка недоступна");
                int count = 0;
                for (int i = 0; results != null && i < results.length(); i++) {
                    String path = results.getJSONObject(i).optString("path");
                    if (path.isEmpty()) continue;
                    File source = new File(path);
                    String name = source.getName();
                    String stem = name.substring(0, name.length() - 5);
                    int suffix = 2;
                    while (folder.findFile(name) != null) name = stem + " (" + suffix++ + ").xlsx";
                    androidx.documentfile.provider.DocumentFile file = folder.createFile(
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", name);
                    if (file == null) throw new IOException("Не удалось создать " + name);
                    try (InputStream in = new FileInputStream(source);
                         OutputStream out = getApplication().getContentResolver().openOutputStream(file.getUri(), "wt")) {
                        if (out == null) throw new IOException("Не удалось записать " + name);
                        copy(in, out);
                    }
                    count++;
                }
                int saved = count;
                main.post(() -> { busy = false; status = "Сохранено Excel-файлов: " + saved; notifyUi(); });
            } catch (Exception e) { fail(e); }
        });
    }

    static void copy(InputStream in, OutputStream out) throws IOException {
        byte[] buffer = new byte[65536]; int n;
        while ((n = in.read(buffer)) != -1) out.write(buffer, 0, n);
    }
    static String readText(InputStream in) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream(); copy(in, out);
        return out.toString(StandardCharsets.UTF_8.name());
    }
    @Override protected void onCleared() { worker.shutdown(); }
}
