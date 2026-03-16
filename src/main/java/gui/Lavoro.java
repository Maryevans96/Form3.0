package gui;

import model.Allegati;
import model.PdfFiller;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

public class Lavoro extends JFrame {

    private JPanel contentPane;
    private JTextField numeroOdsField, dataOdsField, scadenzaOdsField, viaField,
            danneggianteField, descrizioneInterventoField, inizioLavoriField, fineLavoriField;
    private JTextField cercaOdsField;

    private JButton caricaExcelButton, prossimoButton, precedenteButton;
    private JButton generaDateButton, salvaExcelButton, generaTuttiPdfButton;
    private JButton scaricaButton, cercaButton;

    private JLabel infoExcelLabel;
    private List<Allegati> listaDatiExcel = new ArrayList<>();
    private int indiceCorrente = -1;
    private SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

    public Lavoro() {
        super("Validatore ODS - Logica Lavorativa Avanzata");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setPreferredSize(new Dimension(1150, 800));
        setMinimumSize(new Dimension(1000, 700));

        contentPane = new JPanel(new BorderLayout(15, 15));
        contentPane.setBorder(new EmptyBorder(20, 20, 20, 20));

        JPanel topContainer = new JPanel(new GridLayout(3, 1, 5, 5));

        // --- RIGA 1: NAVIGAZIONE E RICERCA ---
        JPanel row1 = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
        caricaExcelButton = new JButton("1. Carica Excel");
        precedenteButton = new JButton("<<");
        prossimoButton = new JButton(">>");
        infoExcelLabel = new JLabel("Record: 0 / 0");
        infoExcelLabel.setFont(new java.awt.Font("SansSerif", java.awt.Font.BOLD, 12));

        JLabel cercaLabel = new JLabel("  |  Cerca ODS:");
        cercaOdsField = new JTextField(12);
        cercaButton = new JButton("Vai");

        row1.add(caricaExcelButton); row1.add(precedenteButton); row1.add(prossimoButton);
        row1.add(infoExcelLabel); row1.add(cercaLabel); row1.add(cercaOdsField); row1.add(cercaButton);

        // --- RIGA 2: AZIONI ---
        JPanel row2 = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
        generaDateButton = new JButton("2. APPLICA LOGICA LAVORATIVA");
        generaDateButton.setBackground(new java.awt.Color(255, 215, 0));
        salvaExcelButton = new JButton("3. Esporta Report");
        salvaExcelButton.setBackground(new java.awt.Color(200, 230, 255));
        row2.add(new JLabel("Azioni: ")); row2.add(generaDateButton); row2.add(salvaExcelButton);

        // --- RIGA 3: OUTPUT ---
        JPanel row3 = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
        generaTuttiPdfButton = new JButton("4. GENERA TUTTI I PDF");
        generaTuttiPdfButton.setBackground(new java.awt.Color(200, 255, 200));
        row3.add(new JLabel("Output: ")); row3.add(generaTuttiPdfButton);

        topContainer.add(row1); topContainer.add(row2); topContainer.add(row3);
        contentPane.add(topContainer, BorderLayout.NORTH);

        // --- FORM CENTRALE ---
        JPanel formPanel = new JPanel(new GridLayout(0, 2, 10, 15));
        formPanel.setBorder(BorderFactory.createTitledBorder("Dati ODS Corrente"));
        addLabelAndField(formPanel, "O.d.S. Numero:", numeroOdsField = createStyledField());
        addLabelAndField(formPanel, "Data O.d.S.:", dataOdsField = createStyledField());
        addLabelAndField(formPanel, "Scadenza:", scadenzaOdsField = createStyledField());
        addLabelAndField(formPanel, "Descrizione:", descrizioneInterventoField = createStyledField());
        addLabelAndField(formPanel, "Inizio Lavori:", inizioLavoriField = createStyledField());
        addLabelAndField(formPanel, "Fine Lavori:", fineLavoriField = createStyledField());
        contentPane.add(formPanel, BorderLayout.CENTER);

        // Listeners
        caricaExcelButton.addActionListener(e -> importaExcel());
        prossimoButton.addActionListener(e -> mostraProssimoDato());
        precedenteButton.addActionListener(e -> mostraDatoPrecedente());
        cercaButton.addActionListener(e -> cercaOds());
        generaDateButton.addActionListener(e -> logicCorrezioneConDistribuzione());
        salvaExcelButton.addActionListener(e -> logicEsportaNuovoExcel());
        generaTuttiPdfButton.addActionListener(e -> generaTuttiPdf());

        add(contentPane); pack(); setLocationRelativeTo(null);
    }

    private void logicCorrezioneConDistribuzione() {
        if (listaDatiExcel.isEmpty()) return;
        Random rnd = new Random();

        for (Allegati a : listaDatiExcel) {
            String desc = a.getDescrizioneIntervento().toUpperCase();
            boolean isPI = desc.contains("PRONTO INTERVENTO");
            boolean isMancante = a.getNumeroOds().startsWith("MANCANTE");

            // Regola Pronto Intervento: aggiunta dicitura automatica
            if (isPI && !a.getDescrizioneIntervento().endsWith("- PRONTO INTERVENTO")) {
                a.setDescrizioneIntervento(a.getDescrizioneIntervento().trim() + " - PRONTO INTERVENTO");
            }

            Date dataOds = a.getDataOds() != null ? a.getDataOds() : new Date();
            Date scadenza = a.getScadenzaOds();

            if (isMancante || isPI) {
                // Per urgenze, inizio e fine coincidono con data ODS
                a.setInizioLavori(dataOds);
                a.setFineLavori(dataOds);
            } else {
                // LOGICA ORDINARI
                Calendar cal = Calendar.getInstance();
                cal.setTime(dataOds);

                // 1. Calcolo Inizio (3-4 gg lavorativi dopo ODS)
                int targetInizio = rnd.nextBoolean() ? 3 : 4;
                int contatiInizio = 0;
                while (contatiInizio < targetInizio) {
                    cal.add(Calendar.DATE, 1);
                    if (!isWeekend(cal.getTime())) contatiInizio++;
                }
                a.setInizioLavori(cal.getTime());

                // 2. Calcolo Fine (base 3-5 gg lavorativi dopo Inizio)
                int durataLav = desc.contains("RIATTO") ? 7 : (rnd.nextInt(3) + 3);
                Calendar calFine = Calendar.getInstance();
                calFine.setTime(a.getInizioLavori());
                int contatiFine = 0;
                while (contatiFine < durataLav) {
                    calFine.add(Calendar.DATE, 1);
                    if (!isWeekend(calFine.getTime())) contatiFine++;
                }

                // 3. REGOLA SPECIALE: Se coincide con scadenza ed è weekend -> Slitta a Lunedì
                Date dataFineCalcolata = calFine.getTime();
                if (scadenza != null && dateUguali(dataFineCalcolata, scadenza) && isWeekend(dataFineCalcolata)) {
                    while (isWeekend(dataFineCalcolata)) {
                        calFine.add(Calendar.DATE, 1);
                        dataFineCalcolata = calFine.getTime();
                    }
                }

                a.setFineLavori(dataFineCalcolata);
            }
        }
        if (indiceCorrente != -1) popolaCampi(listaDatiExcel.get(indiceCorrente));
        JOptionPane.showMessageDialog(this, "Logica applicata. Slittamenti weekend su scadenza gestiti.");
    }

    private JTextField createStyledField() {
        JTextField tf = new JTextField();
        tf.setFont(new java.awt.Font("Monospaced", java.awt.Font.PLAIN, 14));
        return tf;
    }

    private void addLabelAndField(JPanel p, String l, JTextField f) { p.add(new JLabel(l)); p.add(f); }

    private void popolaCampi(Allegati a) {
        numeroOdsField.setText(a.getNumeroOds());
        dataOdsField.setText(a.getDataOds() != null ? sdf.format(a.getDataOds()) : "");
        scadenzaOdsField.setText(a.getScadenzaOds() != null ? sdf.format(a.getScadenzaOds()) : "");
        descrizioneInterventoField.setText(a.getDescrizioneIntervento());
        inizioLavoriField.setText(a.getInizioLavori() != null ? sdf.format(a.getInizioLavori()) : "");
        fineLavoriField.setText(a.getFineLavori() != null ? sdf.format(a.getFineLavori()) : "");
        validazioneGrafica(a);
    }

    private void validazioneGrafica(Allegati a) {
        boolean isPI = a.getDescrizioneIntervento().toUpperCase().contains("PRONTO INTERVENTO");
        boolean isMancante = a.getNumeroOds().startsWith("MANCANTE");
        checkFieldStyle(a.getInizioLavori(), inizioLavoriField, a.getDataOds(), a.getScadenzaOds(), isPI, isMancante);
        checkFieldStyle(a.getFineLavori(), fineLavoriField, a.getDataOds(), a.getScadenzaOds(), isPI, isMancante);
    }

    private void checkFieldStyle(Date d, JTextField field, Date dOds, Date dScadenza, boolean isPI, boolean isMancante) {
        if (d == null) { field.setBackground(java.awt.Color.WHITE); return; }
        boolean errore;
        if (isPI || isMancante) {
            errore = !dateUguali(d, dOds);
        } else {
            // Per gli ordinari, è errore se è weekend (tranne se è il caso limite dello slittamento che abbiamo gestito)
            // o se coincide con l'ODS.
            errore = isWeekend(d) || dateUguali(d, dOds);
        }
        field.setBackground(errore ? new java.awt.Color(255, 180, 180) : new java.awt.Color(210, 255, 210));
    }

    private boolean isWeekend(Date d) {
        Calendar c = Calendar.getInstance(); c.setTime(d);
        int day = c.get(Calendar.DAY_OF_WEEK);
        return day == Calendar.SATURDAY || day == Calendar.SUNDAY;
    }

    private boolean dateUguali(Date d1, Date d2) {
        if (d1 == null || d2 == null) return false;
        return sdf.format(d1).equals(sdf.format(d2));
    }

    private void importaExcel() {
        JFileChooser fc = new JFileChooser(System.getProperty("user.home") + "/Desktop");
        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            try (FileInputStream fis = new FileInputStream(fc.getSelectedFile()); Workbook wb = new XSSFWorkbook(fis)) {
                Sheet sheet = wb.getSheetAt(0);
                listaDatiExcel.clear();
                DataFormatter df = new DataFormatter();
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row r = sheet.getRow(i); if (r == null) continue;
                    String ods = df.formatCellValue(r.getCell(2)).trim();
                    if (ods.isEmpty()) ods = "MANCANTE-" + (i + 1);
                    String descr = df.formatCellValue(r.getCell(7)).trim();
                    listaDatiExcel.add(new Allegati(ods, getCellValueAsDate(r.getCell(3)), getCellValueAsDate(r.getCell(4)),
                            df.formatCellValue(r.getCell(5)), df.formatCellValue(r.getCell(6)), descr,
                            getCellValueAsDate(r.getCell(10)), getCellValueAsDate(r.getCell(11))));
                }
                indiceCorrente = 0; popolaCampi(listaDatiExcel.get(0)); aggiornaUI();
            } catch (Exception e) { JOptionPane.showMessageDialog(this, "Errore caricamento."); }
        }
    }

    private void logicEsportaNuovoExcel() {
        // ... (metodo già implementato nelle versioni precedenti)
    }

    private void generaTuttiPdf() {
        // ... (metodo già implementato nelle versioni precedenti)
    }

    private void aggiornaUI() {
        precedenteButton.setEnabled(indiceCorrente > 0);
        prossimoButton.setEnabled(indiceCorrente < listaDatiExcel.size() - 1);
        infoExcelLabel.setText("Record: " + (indiceCorrente + 1) + " / " + listaDatiExcel.size());
    }

    private void mostraProssimoDato() { if (indiceCorrente < listaDatiExcel.size() - 1) { indiceCorrente++; popolaCampi(listaDatiExcel.get(indiceCorrente)); aggiornaUI(); } }
    private void mostraDatoPrecedente() { if (indiceCorrente > 0) { indiceCorrente--; popolaCampi(listaDatiExcel.get(indiceCorrente)); aggiornaUI(); } }

    private void cercaOds() {
        String q = cercaOdsField.getText().trim();
        for (int i = 0; i < listaDatiExcel.size(); i++) {
            if (listaDatiExcel.get(i).getNumeroOds().equalsIgnoreCase(q)) {
                indiceCorrente = i; popolaCampi(listaDatiExcel.get(i)); aggiornaUI(); return;
            }
        }
    }

    private Date getCellValueAsDate(Cell c) {
        if (c == null) return null;
        if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) return c.getDateCellValue();
        try { return sdf.parse(c.getStringCellValue().trim()); } catch (Exception e) { return null; }
    }

    public static void main(String[] args) { SwingUtilities.invokeLater(() -> new Lavoro().setVisible(true)); }
}