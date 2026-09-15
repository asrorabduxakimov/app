package uz.attestation.excel;

import android.app.AlertDialog;
import android.content.*;
import android.content.res.ColorStateList;
import android.graphics.Color;
import android.graphics.Typeface;
import android.graphics.drawable.GradientDrawable;
import android.net.Uri;
import android.os.Bundle;
import android.text.InputType;
import android.view.*;
import android.widget.*;
import androidx.activity.ComponentActivity;
import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.core.content.FileProvider;
import androidx.core.view.ViewCompat;
import androidx.core.view.WindowInsetsCompat;
import androidx.documentfile.provider.DocumentFile;
import androidx.lifecycle.ViewModelProvider;
import org.json.*;
import java.io.File;
import java.util.*;

public class MainActivity extends ComponentActivity {
    private static final int INK = Color.rgb(24, 48, 40), GREEN = Color.rgb(20, 125, 100);
    private static final String AUTO = "Автоматически";
    private WorkModel model;
    private LinearLayout body, inputList, resultList, rulesList, groupList;
    private TextView status, segments;
    private ProgressBar progress;
    private CheckBox gender, scores;
    private EditText multiplier;
    private Spinner sort, header, prefix;
    private JSONArray ranges = new JSONArray(), groups = new JSONArray();
    private JSONObject displayedInspection;
    private final ArrayList<View> editing = new ArrayList<>(), outputActions = new ArrayList<>();
    private Button run, inspect;

    private final ActivityResultLauncher<Intent> chooseInputs = registerForActivityResult(
        new ActivityResultContracts.StartActivityForResult(), result -> {
            Intent data = result.getData();
            if (result.getResultCode() != RESULT_OK || data == null) return;
            ArrayList<Uri> uris = new ArrayList<>();
            if (data.getClipData() != null) {
                for (int i = 0; i < data.getClipData().getItemCount(); i++) uris.add(data.getClipData().getItemAt(i).getUri());
            } else if (data.getData() != null) uris.add(data.getData());
            model.importFiles(uris);
        });
    private final ActivityResultLauncher<Intent> chooseInputFolder = registerForActivityResult(
        new ActivityResultContracts.StartActivityForResult(), result -> {
            if (result.getResultCode() != RESULT_OK || result.getData() == null) return;
            Uri uri = result.getData().getData();
            if (uri == null) return;
            DocumentFile folder = DocumentFile.fromTreeUri(this, uri);
            ArrayList<Uri> files = new ArrayList<>();
            if (folder != null) {
                for (DocumentFile f : folder.listFiles()) {
                    String name = String.valueOf(f.getName()).toLowerCase(Locale.ROOT);
                    if (f.isFile() && (name.endsWith(".xlsx") || name.endsWith(".xlsm") || name.endsWith(".xls") || name.endsWith(".csv"))) files.add(f.getUri());
                }
            }
            if (files.isEmpty()) message("В выбранной папке нет Excel/CSV-файлов");
            else model.importFiles(files);
        });
    private final ActivityResultLauncher<Intent> saveZip = registerForActivityResult(
        new ActivityResultContracts.StartActivityForResult(), result -> {
            if (result.getResultCode() == RESULT_OK && result.getData() != null && result.getData().getData() != null)
                model.exportZip(result.getData().getData());
        });
    private final ActivityResultLauncher<Intent> saveFiles = registerForActivityResult(
        new ActivityResultContracts.StartActivityForResult(), result -> {
            if (result.getResultCode() == RESULT_OK && result.getData() != null && result.getData().getData() != null)
                model.exportFiles(result.getData().getData());
        });

    @Override public void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        model = new ViewModelProvider(this).get(WorkModel.class);
        buildUi();
        restoreSettings();
        model.changed.observe(this, value -> refresh());
    }

    private void buildUi() {
        ScrollView scroll = new ScrollView(this);
        scroll.setFillViewport(true);
        scroll.setBackgroundColor(Color.rgb(244, 247, 245));
        body = column(); body.setPadding(dp(20), dp(20), dp(20), dp(32));
        scroll.addView(body); setContentView(scroll);
        ViewCompat.setOnApplyWindowInsetsListener(scroll, (view, insets) -> {
            androidx.core.graphics.Insets bars = insets.getInsets(WindowInsetsCompat.Type.systemBars() | WindowInsetsCompat.Type.ime());
            view.setPadding(bars.left, bars.top, bars.right, bars.bottom); return insets;
        });
        body.addView(text("АТТЕСТАЦИЯ", 12, GREEN, true));
        body.addView(text("Ответы → Excel", 30, INK, true));
        body.addView(text("Обработка на телефоне · результаты в ZIP", 14, INK, false));

        LinearLayout input = card("1. Исходные файлы");
        Button add = button(input, "Выбрать файлы", v -> {
            Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT).setType("*/*")
                .addCategory(Intent.CATEGORY_OPENABLE).putExtra(Intent.EXTRA_ALLOW_MULTIPLE, true);
            chooseInputs.launch(intent);
        }); editing.add(add);
        editing.add(button(input, "Выбрать папку", v -> chooseInputFolder.launch(new Intent(Intent.ACTION_OPEN_DOCUMENT_TREE))));
        inputList = column(); input.addView(inputList);

        LinearLayout config = card("2. Настройки обработки");
        inspect = button(config, "Прочитать столбцы и группы первого файла", v -> model.inspect(selected(sort)));
        editing.add(inspect);
        sort = spinner(config, "Столбец сортировки / Savol ID");
        header = spinner(config, "Столбец заголовков вопросов");
        prefix = spinner(config, "Столбец ключа для префиксов");
        gender = new CheckBox(this); gender.setText("Добавить пол по ПИНФЛ"); config.addView(gender); editing.add(gender);
        scores = new CheckBox(this); scores.setText("Добавить баллы и число верных ответов"); config.addView(scores); editing.add(scores);
        config.addView(text("Коэффициент баллов", 14, INK, false));
        multiplier = field("2.0", true); multiplier.setText("2.0"); config.addView(multiplier); editing.add(multiplier);
        scores.setOnCheckedChangeListener((b, checked) -> multiplier.setEnabled(checked && !model.busy));

        config.addView(text("Префиксы по диапазону", 17, INK, true));
        config.addView(text("Например, PM- для ключей от 1 до 10", 13, INK, false));
        rulesList = column(); config.addView(rulesList);
        editing.add(button(config, "+ Диапазон", v -> ruleDialog(false, -1)));
        config.addView(text("Группы Savol ID", 17, INK, true));
        segments = text("Загрузите столбцы, чтобы увидеть размеры групп", 13, INK, false); config.addView(segments);
        groupList = column(); config.addView(groupList);
        editing.add(button(config, "+ Префикс по размеру группы", v -> ruleDialog(true, -1)));
        run = button(body, "Обработать и создать ZIP", v -> startProcessing());
        run.setBackgroundTintList(ColorStateList.valueOf(GREEN)); run.setTextColor(Color.WHITE);
        progress = new ProgressBar(this, null, android.R.attr.progressBarStyleHorizontal);
        body.addView(progress, new LinearLayout.LayoutParams(-1, dp(8)));
        status = text("", 14, INK, false); status.setTextIsSelectable(true); body.addView(status);

        LinearLayout results = card("3. Результаты");
        results.addView(text("Выберите Telegram в меню отправки, затем чат и подтвердите отправку.", 14, INK, false));
        outputActions.add(button(results, "Сохранить ZIP", v -> saveZip.launch(
            new Intent(Intent.ACTION_CREATE_DOCUMENT).addCategory(Intent.CATEGORY_OPENABLE)
                .setType("application/zip").putExtra(Intent.EXTRA_TITLE, "Attestation_results.zip"))));
        outputActions.add(button(results, "Отправить ZIP → Telegram", v -> share(true, null)));
        outputActions.add(button(results, "Отправить все Excel-файлы → Telegram", v -> share(false, null)));
        outputActions.add(button(results, "Сохранить Excel-файлы в папку", v -> saveFiles.launch(new Intent(Intent.ACTION_OPEN_DOCUMENT_TREE))));
        resultList = column(); results.addView(resultList);
        button(results, "Скопировать число участников", v -> {
            JSONArray rows = model.report.optJSONArray("results");
            if (rows == null) { message("Сначала обработайте файлы"); return; }
            StringBuilder counts = new StringBuilder();
            for (int i = 0; i < rows.length(); i++) {
                JSONObject r = rows.optJSONObject(i);
                counts.append(r.optString("name")).append('\t').append(r.optString("error").isEmpty() ? r.optInt("participants") : "ОШИБКА").append('\n');
            }
            ((ClipboardManager) getSystemService(CLIPBOARD_SERVICE)).setPrimaryClip(ClipData.newPlainText("Участники", counts.toString()));
            message("Список скопирован");
        });
        button(results, "Журнал обработки", v -> {
            TextView log = text(model.report.optString("log", "Обработка ещё не выполнялась"), 12, INK, false);
            log.setTextIsSelectable(true); log.setPadding(dp(16), dp(8), dp(16), dp(8));
            ScrollView logs = new ScrollView(this); logs.addView(log);
            new AlertDialog.Builder(this).setTitle("Журнал").setView(logs).setPositiveButton("Закрыть", null).show();
        });
    }

    private void refresh() {
        for (View v : editing) v.setEnabled(!model.busy);
        inspect.setEnabled(!model.busy && !model.inputs.isEmpty());
        multiplier.setEnabled(!model.busy && scores.isChecked());
        run.setEnabled(!model.busy && !model.inputs.isEmpty());
        run.setText(model.busy ? "Пожалуйста, подождите…" : "Обработать и создать ZIP");
        progress.setProgress(model.progress); status.setText(model.status);
        for (View v : outputActions) v.setEnabled(!model.busy && !model.report.optString("archive").isEmpty());
        inputList.removeAllViews();
        if (model.inputs.isEmpty()) inputList.addView(text("Файлы пока не выбраны", 14, INK, false));
        for (JSONObject item : model.inputs) {
            LinearLayout row = column(); inputList.addView(row);
            row.addView(text(item.optString("name"), 14, INK, true));
            Button remove = button(row, "Убрать из списка", v -> { model.inputs.remove(item); model.saveInputs(); });
            remove.setEnabled(!model.busy);
        }
        if (displayedInspection != model.inspection && model.inspection.has("headers")) {
            displayedInspection = model.inspection;
            for (Spinner s : new Spinner[]{sort, header, prefix}) setHeaders(s, model.inspection.optJSONArray("headers"), selected(s));
            StringBuilder summary = new StringBuilder("Столбец: " + model.inspection.optString("sort_column") + "\n");
            JSONArray found = model.inspection.optJSONArray("segments");
            if (found == null || found.length() == 0) summary.append("Числовые группы не найдены");
            else for (int i = 0; i < found.length(); i++) {
                JSONObject g = found.optJSONObject(i);
                summary.append(g.optInt("start")).append("–").append(g.optInt("end"))
                    .append(": ").append(g.optInt("count")).append(" вопросов\n");
            }
            segments.setText(summary);
        }
        renderRules();
        resultList.removeAllViews();
        JSONArray results = model.report.optJSONArray("results");
        for (int i = 0; results != null && i < results.length(); i++) {
            JSONObject r = results.optJSONObject(i);
            String error = r.optString("error"), path = r.optString("path");
            resultList.addView(text(r.optString("name"), 15, INK, true));
            String detail = error.isEmpty() ? "Участников: " + r.optInt("participants")
                + (r.optInt("error_rows") > 0 ? " · строк на листе «Ошибки»: " + r.optInt("error_rows") : "") : "Ошибка: " + error;
            resultList.addView(text(detail, 13, error.isEmpty() ? GREEN : Color.rgb(165, 45, 45), false));
            if (!path.isEmpty()) {
                Button send = button(resultList, "Отправить этот файл", v -> share(false, path));
                send.setEnabled(!model.busy);
            }
        }
    }

    private void startProcessing() {
        try {
            double factor = Double.parseDouble(multiplier.getText().toString().trim().replace(',', '.'));
            if (!Double.isFinite(factor) || factor < 0) throw new IllegalArgumentException("Проверьте коэффициент баллов");
            JSONObject options = new JSONObject().put("add_gender", gender.isChecked()).put("add_scores", scores.isChecked())
                .put("multiplier", factor).put("sort_column", selected(sort)).put("header_column", selected(header))
                .put("prefix_column", selected(prefix)).put("suffixes", ranges).put("size_rules", groups);
            saveSettings(); model.process(options);
        } catch (Exception e) { message("Проверьте настройки: " + e.getMessage()); }
    }

    private void ruleDialog(boolean group, int index) {
        JSONArray collection = group ? groups : ranges;
        JSONObject current = index >= 0 ? collection.optJSONObject(index) : new JSONObject();
        LinearLayout fields = column(); fields.setPadding(dp(20), dp(4), dp(20), dp(4));
        fields.addView(text("Префикс", 14, INK, false));
        EditText value = field("PM-", false); value.setText(current.optString("prefix")); fields.addView(value);
        fields.addView(text(group ? "Количество вопросов в группе" : "От (включительно)", 14, INK, false));
        EditText from = field(group ? "10" : "1", true); fields.addView(from);
        from.setText(current.optString(group ? "count" : "from", ""));
        EditText to = field("10", true);
        if (!group) { fields.addView(text("До (включительно)", 14, INK, false)); fields.addView(to); to.setText(current.optString("to", "")); }
        AlertDialog dialog = new AlertDialog.Builder(this).setTitle(group ? "Префикс группы" : "Префикс диапазона")
            .setView(fields).setPositiveButton("Сохранить", null).setNegativeButton("Отмена", null).create();
        dialog.setOnShowListener(d -> dialog.getButton(AlertDialog.BUTTON_POSITIVE).setOnClickListener(v -> {
            try {
                String p = value.getText().toString().trim();
                if (p.isEmpty()) throw new IllegalArgumentException("Введите префикс");
                JSONObject rule = new JSONObject().put("prefix", p);
                if (group) {
                    int count = Integer.parseInt(from.getText().toString().trim());
                    if (count <= 0) throw new IllegalArgumentException("Размер должен быть больше нуля");
                    for (int i = 0; i < collection.length(); i++)
                        if (i != index && collection.getJSONObject(i).optInt("count") == count)
                            throw new IllegalArgumentException("Правило для этого размера уже есть");
                    rule.put("count", count);
                } else {
                    double lo = Double.parseDouble(from.getText().toString().trim().replace(',', '.'));
                    double hi = Double.parseDouble(to.getText().toString().trim().replace(',', '.'));
                    if (!Double.isFinite(lo) || !Double.isFinite(hi) || lo > hi) throw new IllegalArgumentException("Проверьте границы диапазона");
                    rule.put("from", lo).put("to", hi);
                }
                if (index < 0) collection.put(rule); else collection.put(index, rule);
                renderRules(); saveSettings(); dialog.dismiss();
            } catch (Exception e) { message("Не удалось сохранить: " + e.getMessage()); }
        }));
        dialog.show();
    }

    private void renderRules() {
        renderRuleList(rulesList, ranges, false); renderRuleList(groupList, groups, true);
    }
    private void renderRuleList(LinearLayout parent, JSONArray values, boolean group) {
        parent.removeAllViews();
        for (int i = 0; i < values.length(); i++) {
            int index = i; JSONObject rule = values.optJSONObject(i);
            String label = rule.optString("prefix") + " · " + (group ? rule.optInt("count") + " вопросов"
                : rule.optString("from") + " … " + rule.optString("to"));
            Button edit = button(parent, label, v -> ruleDialog(group, index)); edit.setEnabled(!model.busy);
            Button remove = button(parent, "Удалить правило", v -> { values.remove(index); renderRules(); saveSettings(); }); remove.setEnabled(!model.busy);
        }
    }

    private void share(boolean archive, String single) {
        try {
            ArrayList<Uri> uris = new ArrayList<>();
            if (archive) uris.add(fileUri(model.report.getString("archive")));
            else if (single != null) uris.add(fileUri(single));
            else {
                JSONArray results = model.report.getJSONArray("results");
                for (int i = 0; i < results.length(); i++) {
                    String path = results.getJSONObject(i).optString("path");
                    if (!path.isEmpty()) uris.add(fileUri(path));
                }
            }
            if (uris.isEmpty()) { message("Нет файлов для отправки"); return; }
            Intent intent = new Intent(uris.size() == 1 ? Intent.ACTION_SEND : Intent.ACTION_SEND_MULTIPLE);
            intent.setType(archive ? "application/zip" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            if (uris.size() == 1) intent.putExtra(Intent.EXTRA_STREAM, uris.get(0));
            else intent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);
            ClipData clips = ClipData.newRawUri("Результаты аттестации", uris.get(0));
            for (int i = 1; i < uris.size(); i++) clips.addItem(new ClipData.Item(uris.get(i)));
            intent.setClipData(clips); intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
            startActivity(Intent.createChooser(intent, "Отправить через Telegram"));
        } catch (Exception e) { message("Не удалось открыть отправку: " + e.getMessage()); }
    }
    private Uri fileUri(String path) {
        File file = new File(path);
        if (!file.isFile()) throw new IllegalArgumentException("Файл результата отсутствует. Повторите обработку.");
        return FileProvider.getUriForFile(this, getPackageName() + ".files", file);
    }

    private void restoreSettings() {
        SharedPreferences p = getSharedPreferences("app", 0);
        gender.setChecked(p.getBoolean("gender", false)); scores.setChecked(p.getBoolean("scores", false));
        multiplier.setText(p.getString("multiplier", "2.0"));
        try {
            ranges = new JSONArray(p.getString("ranges", "[]")); groups = new JSONArray(p.getString("groups", "[]"));
            JSONArray headers = new JSONArray(p.getString("headers", "[]"));
            setHeaders(sort, headers, p.getString("sort", ""));
            setHeaders(header, headers, p.getString("header", ""));
            setHeaders(prefix, headers, p.getString("prefix", ""));
        } catch (JSONException e) { message("Настройки восстановлены частично"); }
    }
    private void saveSettings() {
        if (sort == null) return;
        JSONArray headers = new JSONArray();
        for (int i = 1; i < sort.getCount(); i++) headers.put(String.valueOf(sort.getItemAtPosition(i)));
        getSharedPreferences("app", 0).edit().putBoolean("gender", gender.isChecked()).putBoolean("scores", scores.isChecked())
            .putString("multiplier", multiplier.getText().toString()).putString("sort", selected(sort))
            .putString("header", selected(header)).putString("prefix", selected(prefix))
            .putString("ranges", ranges.toString()).putString("groups", groups.toString()).putString("headers", headers.toString()).apply();
    }
    @Override protected void onPause() { saveSettings(); super.onPause(); }

    private Spinner spinner(LinearLayout parent, String label) {
        parent.addView(text(label, 14, INK, false)); Spinner spinner = new Spinner(this);
        setHeaders(spinner, new JSONArray(), ""); parent.addView(spinner, new LinearLayout.LayoutParams(-1, dp(52)));
        editing.add(spinner); return spinner;
    }
    private void setHeaders(Spinner spinner, JSONArray headers, String previous) {
        ArrayList<String> values = new ArrayList<>(); values.add(AUTO);
        if (headers != null) for (int i = 0; i < headers.length(); i++) values.add(headers.optString(i));
        ArrayAdapter<String> adapter = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, values);
        adapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
        spinner.setAdapter(adapter); spinner.setSelection(Math.max(0, values.indexOf(previous)));
    }
    private String selected(Spinner spinner) { return spinner.getSelectedItemPosition() <= 0 ? "" : String.valueOf(spinner.getSelectedItem()); }
    private LinearLayout column() { LinearLayout layout = new LinearLayout(this); layout.setOrientation(LinearLayout.VERTICAL); return layout; }
    private LinearLayout card(String title) {
        LinearLayout card = column(); card.setPadding(dp(16), dp(14), dp(16), dp(16));
        GradientDrawable bg = new GradientDrawable(); bg.setColor(Color.WHITE); bg.setCornerRadius(dp(18)); card.setBackground(bg);
        LinearLayout.LayoutParams p = new LinearLayout.LayoutParams(-1, -2); p.setMargins(0, dp(18), 0, dp(8));
        body.addView(card, p); card.addView(text(title, 19, INK, true)); return card;
    }
    private TextView text(String value, int size, int color, boolean bold) {
        TextView view = new TextView(this); view.setText(value); view.setTextSize(size); view.setTextColor(color);
        view.setPadding(0, dp(7), 0, dp(7)); if (bold) view.setTypeface(Typeface.DEFAULT, Typeface.BOLD); return view;
    }
    private Button button(LinearLayout parent, String label, View.OnClickListener listener) {
        Button b = new Button(this); b.setText(label); b.setAllCaps(false); b.setTextSize(14); b.setMinHeight(dp(48));
        b.setOnClickListener(listener); parent.addView(b, new LinearLayout.LayoutParams(-1, -2)); return b;
    }
    private EditText field(String hint, boolean numeric) {
        EditText e = new EditText(this); e.setHint(hint); e.setSingleLine(true); e.setTextSize(16);
        e.setInputType(numeric ? InputType.TYPE_CLASS_NUMBER | InputType.TYPE_NUMBER_FLAG_DECIMAL | InputType.TYPE_NUMBER_FLAG_SIGNED : InputType.TYPE_CLASS_TEXT);
        e.setMinHeight(dp(48)); return e;
    }
    private int dp(int value) { return Math.round(value * getResources().getDisplayMetrics().density); }
    private void message(String value) { Toast.makeText(this, value, Toast.LENGTH_LONG).show(); }
}
