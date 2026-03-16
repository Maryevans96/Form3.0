package gui;

import model.Allegati;
import model.PdfFiller;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

public class Lavoro extends JFrame {

    private JPanel contentPane;
    private JTextField numeroOdsField, dataOdsField, scadenzaOdsField, viaField,
            danneggianteField, descrizioneInterventoField, inizioLavoriField, fineLavoriField;

    private JButton caricaExcelButton, prossimoButton, precedenteButton, randomDateButton, randomAllButton, generaTuttiPdfButton;
    private JButton compilaButton;
    private JButton pulisciCampiButton;

    private JLabel infoExcelLabel;

    private List<Allegati> listaDatiExcel = new ArrayList<>();
    private int indiceCorrente = -1;
    private String currentExcelPath;
    private SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

    public Lavoro() {
        super("Automazione Completa PDF & Excel");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setPreferredSize(new Dimension(1100, 800));

        contentPane = new JPanel(new BorderLayout(15, 15));
        contentPane.setBorder(new EmptyBorder(20, 20, 20, 20));

        // --- Pannello Comandi Superiori ---
        JPanel topContainer = new JPanel(new GridLayout(3, 1, 5, 5));

        JPanel row1 = new JPanel(new FlowLayout(FlowLayout.LEFT));
        caricaExcelButton = new JButton("1. Carica Excel");
        infoExcelLabel = new JLabel("Nessun file caricato");
        precedenteButton = new JButton("<<");
        prossimoButton = new JButton(">>");
        row1.add(caricaExcelButton);
        row1.add(precedenteButton);
        row1.add(prossimoButton);
        row1.add(infoExcelLabel);

        JPanel row2 = new JPanel(new FlowLayout(FlowLayout.LEFT));
        randomDateButton = new JButton("Genera Date (Singolo)");
        randomAllButton = new JButton("Genera Date (TUTTI)");
        randomAllButton.setBackground(new Color(200, 230, 255));
        row2.add(new JLabel("Excel: "));
        row2.add(randomDateButton);
        row2.add(randomAllButton);

        JPanel row3 = new JPanel(new FlowLayout(FlowLayout.LEFT));
        generaTuttiPdfButton = new JButton("GENERA TUTTI I PDF");
        generaTuttiPdfButton.setBackground(new Color(200, 255, 200));
        generaTuttiPdfButton.setFont(new Font("SansSerif", Font.BOLD, 12));

        compilaButton = new JButton("Compila");
        pulisciCampiButton = new JButton("Pulisci");

        row3.add(new JLabel("PDF: "));
        row3.add(generaTuttiPdfButton);

        topContainer.add(row1);
        topContainer.add(row2);
        topContainer.add(row3);
        contentPane.add(topContainer, BorderLayout.NORTH);

        // --- Pannello Centrale (Form) ---
        JPanel formPanel = new JPanel(new GridLayout(0, 2, 10, 15));
        formPanel.setBorder(BorderFactory.createTitledBorder("Anteprima Dati"));

        numeroOdsField = new JTextField();
        dataOdsField = new JTextField();
        scadenzaOdsField = new JTextField();
        viaField = new JTextField();
        danneggianteField = new JTextField();
        descrizioneInterventoField = new JTextField();
        inizioLavoriField = new JTextField();
        fineLavoriField = new JTextField();

        addLabelAndField(formPanel, "O.d.S.:", numeroOdsField);
        addLabelAndField(formPanel, "Data O.d.S.:", dataOdsField);
        addLabelAndField(formPanel, "Scadenza:", scadenzaOdsField);
        addLabelAndField(formPanel, "Via:", viaField);
        addLabelAndField(formPanel, "Danneggiante:", danneggianteField);
        addLabelAndField(formPanel, "Descrizione:", descrizioneInterventoField);
        addLabelAndField(formPanel, "Inizio Lavori:", inizioLavoriField);
        addLabelAndField(formPanel, "Fine Lavori:", fineLavoriField);

        contentPane.add(formPanel, BorderLayout.CENTER);

        // --- Eventi ---
        caricaExcelButton.addActionListener(e -> importaExcel());
        prossimoButton.addActionListener(e -> mostraProssimoDato());
        precedenteButton.addActionListener(e -> mostraDatoPrecedente());
        randomDateButton.addActionListener(e -> generaESalvaSingolo());
        randomAllButton.addActionListener(e -> generaESalvaTutti());
        generaTuttiPdfButton.addActionListener(e -> generaTuttiPdf());
        pulisciCampiButton.addActionListener(e -> pulisciCampi());

        add(contentPane);
        pack();
        setLocationRelativeTo(null);
    }

    private void pulisciCampi() {
        numeroOdsField.setText("");
        dataOdsField.setText("");
        scadenzaOdsField.setText("");
        viaField.setText("");
        danneggianteField.setText("");
        descrizioneInterventoField.setText("");
        inizioLavoriField.setText("");
        fineLavoriField.setText("");
    }

    private void generaTuttiPdf() {
        if (listaDatiExcel.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Carica prima un file Excel!");
            return;
        }
        String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        File cartellaDest = new File(System.getProperty("user.home") + "/Desktop/PDF_GENERATI_" + timeStamp);
        if (!cartellaDest.exists()) cartellaDest.mkdirs();

        int prodotti = 0;
        PdfFiller filler = new PdfFiller();
        try {
            for (Allegati a : listaDatiExcel) {
                if (a.getNumeroOds() != null && !a.getNumeroOds().isEmpty()) {
                    String nomeFile = a.getNumeroOds().replaceAll("[\\\\/:*?\"<>|]", "_") + ".pdf";
                    File outputPdf = new File(cartellaDest, nomeFile);
                    filler.fillPdfSpecificFields(outputPdf.getAbsolutePath(), a);
                    prodotti++;
                }
            }
            JOptionPane.showMessageDialog(this, "Completato! " + prodotti + " PDF generati.");
            Desktop.getDesktop().open(cartellaDest);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Errore generazione PDF: " + ex.getMessage());
        }
    }

    private void generaESalvaTutti() {
        if (listaDatiExcel.isEmpty()) return;
        Map<String, Date[]> mappaDate = new HashMap<>();
        for (Allegati a : listaDatiExcel) {
            Date[] date = calcolaDateLogiche(a);
            if (date != null) mappaDate.put(a.getNumeroOds(), date);
        }
        salvaDateSuExcel(mappaDate);
        importaExcelSilenzioso();
        JOptionPane.showMessageDialog(this, "Date aggiornate per tutti i record!");
    }

    private void generaESalvaSingolo() {
        if (indiceCorrente == -1) return;
        Allegati a = listaDatiExcel.get(indiceCorrente);
        Date[] date = calcolaDateLogiche(a);
        if (date != null) {
            Map<String, Date[]> singola = new HashMap<>();
            singola.put(a.getNumeroOds(), date);
            salvaDateSuExcel(singola);
            importaExcelSilenzioso();
        }
    }

    // NUOVA LOGICA DI CALCOLO
    private Date[] calcolaDateLogiche(Allegati a) {
        if (a.getDataOds() == null) return null;

        String desc = a.getDescrizioneIntervento() != null ? a.getDescrizioneIntervento().toUpperCase() : "";

        // REGOLA 1: PRONTO INTERVENTO -> Date uguali all'ODS
        if (desc.contains("PRONTO INTERVENTO")) {
            return new Date[]{ a.getDataOds(), a.getDataOds() };
        }

        // REGOLA 2: RIATTO ALLOGGIO -> Durata 7 giorni, altrimenti 3/4
        int durata = desc.contains("RIATTO ALLOGGIO") ? 7 : (new Random().nextBoolean() ? 3 : 4);

        return calcolaFinestraRandom(a.getDataOds(), a.getScadenzaOds(), durata);
    }

    private Date[] calcolaFinestraRandom(Date dataOds, Date dataScadenza, int durata) {
        if (dataScadenza == null) return null;
        List<Date> feriali = new ArrayList<>();
        Calendar cal = Calendar.getInstance();
        cal.setTime(dataOds);
        cal.add(Calendar.DAY_OF_MONTH, 1);

        while (cal.getTime().before(dataScadenza)) {
            int dow = cal.get(Calendar.DAY_OF_WEEK);
            if (dow != Calendar.SATURDAY && dow != Calendar.SUNDAY) feriali.add(cal.getTime());
            cal.add(Calendar.DAY_OF_MONTH, 1);
        }

        if (feriali.size() < durata) durata = feriali.size();
        if (feriali.isEmpty()) return null;

        int start = new Random().nextInt(feriali.size() - durata + 1);
        return new Date[]{ feriali.get(start), feriali.get(start + durata - 1) };
    }

    private void salvaDateSuExcel(Map<String, Date[]> mappaDate) {
        try (FileInputStream fis = new FileInputStream(currentExcelPath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            CellStyle ds = workbook.createCellStyle();
            ds.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd/mm/yyyy"));
            for (Row row : sheet) {
                String ods = new DataFormatter().formatCellValue(row.getCell(2)).trim();
                if (mappaDate.containsKey(ods)) {
                    Date[] d = mappaDate.get(ods);
                    updateCell(row, 10, d[0], ds);
                    updateCell(row, 11, d[1], ds);
                }
            }
            try (FileOutputStream fos = new FileOutputStream(currentExcelPath)) { workbook.write(fos); }
        } catch (Exception e) { JOptionPane.showMessageDialog(this, "Errore Excel: " + e.getMessage()); }
    }

    private void updateCell(Row r, int col, Date val, CellStyle s) {
        Cell c = r.getCell(col);
        if (c == null) c = r.createCell(col);
        c.setCellValue(val);
        c.setCellStyle(s);
    }

    private void importaExcel() {
        JFileChooser fc = new JFileChooser(System.getProperty("user.home") + "/Desktop");
        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            currentExcelPath = fc.getSelectedFile().getAbsolutePath();
            importaExcelSilenzioso();
        }
    }

    private void importaExcelSilenzioso() {
        try (FileInputStream fis = new FileInputStream(currentExcelPath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            listaDatiExcel.clear();
            DataFormatter df = new DataFormatter();

            // Partiamo da i=1 per saltare l'intestazione
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row r = sheet.getRow(i);
                if (r == null) continue;

                // Leggiamo il numero ODS (Colonna C - indice 2)
                String ods = df.formatCellValue(r.getCell(2)).trim();

                // L'UNICO motivo per scartare una riga deve essere l'assenza del numero ODS
                if (ods.isEmpty()) continue;

                // Leggiamo la descrizione (Colonna H - indice 7)
                String descrizione = df.formatCellValue(r.getCell(7));

                // Aggiungiamo alla lista SENZA FILTRARE i Pronto Intervento
                listaDatiExcel.add(new Allegati(
                        ods,
                        getCellValueAsDate(r.getCell(3)),  // Data ODS
                        getCellValueAsDate(r.getCell(4)),  // Scadenza
                        df.formatCellValue(r.getCell(5)),  // Via
                        df.formatCellValue(r.getCell(6)),  // Danneggiante
                        descrizione,                       // Descrizione (qui legge PRONTO INTERVENTO o RIATTO)
                        getCellValueAsDate(r.getCell(10)), // Data Inizio esistente
                        getCellValueAsDate(r.getCell(11))  // Data Fine esistente
                ));
            }

            // Reset indice se necessario e aggiornamento UI
            if (indiceCorrente == -1 && !listaDatiExcel.isEmpty()) indiceCorrente = 0;
            if (indiceCorrente >= listaDatiExcel.size()) indiceCorrente = listaDatiExcel.size() - 1;

            if (indiceCorrente != -1) popolaCampi(listaDatiExcel.get(indiceCorrente));

            infoExcelLabel.setText("Record: " + (indiceCorrente + 1) + " / " + listaDatiExcel.size());

        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Errore caricamento: " + e.getMessage());
        }
    }

    private void mostraProssimoDato() {
        if (indiceCorrente < listaDatiExcel.size() - 1) {
            indiceCorrente++;
            popolaCampi(listaDatiExcel.get(indiceCorrente));
            infoExcelLabel.setText("Record: " + (indiceCorrente + 1) + " / " + listaDatiExcel.size());
        }
    }

    private void mostraDatoPrecedente() {
        if (indiceCorrente > 0) {
            indiceCorrente--;
            popolaCampi(listaDatiExcel.get(indiceCorrente));
            infoExcelLabel.setText("Record: " + (indiceCorrente + 1) + " / " + listaDatiExcel.size());
        }
    }

    private void addLabelAndField(JPanel p, String l, JTextField f) {
        p.add(new JLabel(l) {{ setFont(new Font("SansSerif", Font.BOLD, 12)); }});
        p.add(f);
    }

    private void popolaCampi(Allegati a) {
        numeroOdsField.setText(a.getNumeroOds());
        dataOdsField.setText(a.getDataOds() != null ? sdf.format(a.getDataOds()) : "");
        scadenzaOdsField.setText(a.getScadenzaOds() != null ? sdf.format(a.getScadenzaOds()) : "");
        viaField.setText(a.getVia());
        danneggianteField.setText(a.getDanneggiante());
        descrizioneInterventoField.setText(a.getDescrizioneIntervento());

        inizioLavoriField.setText(a.getInizioLavori() != null ? sdf.format(a.getInizioLavori()) : "");
        fineLavoriField.setText(a.getFineLavori() != null ? sdf.format(a.getFineLavori()) : "");
    }

    private Date getCellValueAsDate(Cell c) {
        if (c == null) return null;
        // Se la cella è formattata come DATA in Excel
        if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) {
            return c.getDateCellValue();
        }
        // Se la cella è TESTO (es. "15/03/2024")
        if (c.getCellType() == CellType.STRING) {
            try {
                return sdf.parse(c.getStringCellValue().trim());
            } catch (Exception e) {
                return null;
            }
        }
        return null;
    }

    public static void main(String[] args) { SwingUtilities.invokeLater(() -> new Lavoro().setVisible(true)); }
}