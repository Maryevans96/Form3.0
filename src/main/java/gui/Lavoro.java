package gui;

import model.Allegati;
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
    private JTextField numeroOdsField, dataOdsField, scadenzaOdsField, descrizioneInterventoField, inizioLavoriField, fineLavoriField, cercaOdsField;
    private JButton caricaExcelButton, prossimoButton, precedenteButton, generaDateButton, salvaExcelButton, generaTuttiPdfButton, cercaButton, verificaDateMancantiButton;
    private JButton salvaModificheManualiButton, refreshButton, trovaErroreButton;

    private JLabel infoExcelLabel;
    private List<Allegati> listaDatiExcel = new ArrayList<>();
    private Set<Integer> recordSalvatiManualmente = new HashSet<>();
    private int indiceCorrente = -1;
    private SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

    public Lavoro() {
        super("Validatore ODS - Logica Lavorativa Avanzata");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setPreferredSize(new Dimension(1250, 800));
        setMinimumSize(new Dimension(1000, 700));

        contentPane = new JPanel(new BorderLayout(15, 15));
        contentPane.setBorder(new EmptyBorder(20, 20, 20, 20));

        JPanel topContainer = new JPanel(new GridLayout(3, 1, 5, 5));

        // --- RIGA 1: NAVIGAZIONE ---
        JPanel row1 = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
        caricaExcelButton = new JButton("1. Carica Excel");
        precedenteButton = new JButton("<<");
        prossimoButton = new JButton(">>");
        infoExcelLabel = new JLabel("Record: 0 / 0");
        cercaOdsField = new JTextField(12);
        cercaButton = new JButton("Vai");

        trovaErroreButton = new JButton("Trova Errore");
        trovaErroreButton.setBackground(new java.awt.Color(255, 102, 102));
        trovaErroreButton.setForeground(java.awt.Color.WHITE);

        row1.add(caricaExcelButton); row1.add(precedenteButton); row1.add(prossimoButton);
        row1.add(infoExcelLabel); row1.add(new JLabel(" | ")); row1.add(trovaErroreButton);
        row1.add(new JLabel(" | Cerca:")); row1.add(cercaOdsField); row1.add(cercaButton);

        // --- RIGA 2: AZIONI ---
        JPanel row2 = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
        generaDateButton = new JButton("Logica Standard (Reset)");
        generaDateButton.setBackground(new java.awt.Color(255, 215, 0));

        verificaDateMancantiButton = new JButton("Genera Mancanti (3-4gg)");
        verificaDateMancantiButton.setBackground(new java.awt.Color(180, 255, 180));

        salvaModificheManualiButton = new JButton("SALVA MODIFICHE");
        salvaModificheManualiButton.setBackground(new java.awt.Color(100, 149, 237));
        salvaModificheManualiButton.setForeground(java.awt.Color.WHITE);

        refreshButton = new JButton("Refresh");
        salvaExcelButton = new JButton("Esporta Excel");

        row2.add(generaDateButton); row2.add(verificaDateMancantiButton); row2.add(new JLabel(" | "));
        row2.add(salvaModificheManualiButton); row2.add(refreshButton); row2.add(salvaExcelButton);

        // --- RIGA 3: OUTPUT ---
        JPanel row3 = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 5));
        generaTuttiPdfButton = new JButton("4. GENERA TUTTI I PDF");
        row3.add(new JLabel("Output: ")); row3.add(generaTuttiPdfButton);

        topContainer.add(row1); topContainer.add(row2); topContainer.add(row3);
        contentPane.add(topContainer, BorderLayout.NORTH);

        // FORM CENTRALE
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
        verificaDateMancantiButton.addActionListener(e -> logicGeneraDateIntervalloRandom());
        salvaModificheManualiButton.addActionListener(e -> salvaDatiManuali());
        refreshButton.addActionListener(e -> { if (indiceCorrente != -1) popolaCampi(listaDatiExcel.get(indiceCorrente)); });
        trovaErroreButton.addActionListener(e -> vaiAlProssimoErrore());

        add(contentPane); pack(); setLocationRelativeTo(null);
    }

    // --- LOGICA PULSANTI ---

    private void logicCorrezioneConDistribuzione() {
        if (listaDatiExcel.isEmpty()) return;
        Random rnd = new Random();
        int contatore = 0;
        for (Allegati a : listaDatiExcel) {
            eseguiLogicaSuRecord(a, rnd);
            contatore++;
        }
        if (indiceCorrente != -1) popolaCampi(listaDatiExcel.get(indiceCorrente));
        JOptionPane.showMessageDialog(this, "LOGICA STANDARD: Rielaborati tutti i " + contatore + " record.");
    }

    private void logicGeneraDateIntervalloRandom() {
        if (listaDatiExcel.isEmpty()) return;
        Random rnd = new Random();
        int contatore = 0;
        for (Allegati a : listaDatiExcel) {
            if (a.getInizioLavori() == null || a.getFineLavori() == null) {
                eseguiLogicaSuRecord(a, rnd);
                contatore++;
            }
        }
        if (indiceCorrente != -1) popolaCampi(listaDatiExcel.get(indiceCorrente));
        JOptionPane.showMessageDialog(this, "GENERA MANCANTI: Popolati " + contatore + " record vuoti.");
    }

    private void eseguiLogicaSuRecord(Allegati a, Random rnd) {
        String desc = a.getDescrizioneIntervento().toUpperCase();
        boolean isPI = desc.contains("PRONTO INTERVENTO");

        // Regola Suffisso Descrittivo
        if (isPI && !a.getDescrizioneIntervento().endsWith("- PRONTO INTERVENTO")) {
            a.setDescrizioneIntervento(a.getDescrizioneIntervento().trim() + " - PRONTO INTERVENTO");
        }

        if (isPI) {
            a.setInizioLavori(a.getDataOds());
            a.setFineLavori(a.getDataOds());
        } else {
            Date dOds = a.getDataOds();
            Date dScadenza = a.getScadenzaOds();
            if (dOds != null) {
                Calendar calInizio = Calendar.getInstance();
                calInizio.setTime(dOds);
                do { calInizio.add(Calendar.DATE, 1); } while (isWeekend(calInizio.getTime()));

                int durataTarget = rnd.nextBoolean() ? 3 : 4;
                Calendar calFine = (Calendar) calInizio.clone();
                int lavorativiContati = 0;
                while (lavorativiContati < durataTarget) {
                    calFine.add(Calendar.DATE, 1);
                    if (!isWeekend(calFine.getTime())) lavorativiContati++;
                }
                while (isWeekend(calFine.getTime())) { calFine.add(Calendar.DATE, 1); }

                if (dScadenza != null && !calFine.getTime().before(dScadenza)) {
                    calFine.setTime(dScadenza);
                    calFine.add(Calendar.DATE, -1);
                    while (isWeekend(calFine.getTime())) calFine.add(Calendar.DATE, -1);
                }
                a.setInizioLavori(calInizio.getTime());
                a.setFineLavori(calFine.getTime());
            }
        }
    }

    private void salvaDatiManuali() {
        if (indiceCorrente == -1 || listaDatiExcel.isEmpty()) return;
        try {
            Allegati a = listaDatiExcel.get(indiceCorrente);
            a.setNumeroOds(numeroOdsField.getText().trim());

            String desc = descrizioneInterventoField.getText().trim();
            if (desc.toUpperCase().contains("PRONTO INTERVENTO") && !desc.endsWith("- PRONTO INTERVENTO")) {
                desc += " - PRONTO INTERVENTO";
            }
            a.setDescrizioneIntervento(desc);

            a.setDataOds(parseData(dataOdsField.getText()));
            a.setScadenzaOds(parseData(scadenzaOdsField.getText()));
            a.setInizioLavori(parseData(inizioLavoriField.getText()));
            a.setFineLavori(parseData(fineLavoriField.getText()));

            recordSalvatiManualmente.add(indiceCorrente);
            popolaCampi(a);
            JOptionPane.showMessageDialog(this, "Record salvato e confermato (Verde).");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Errore formato data!", "Errore", JOptionPane.ERROR_MESSAGE);
        }
    }

    // --- NAVIGAZIONE E UI ---

    private void vaiAlProssimoErrore() {
        if (listaDatiExcel.isEmpty()) return;
        int partenza = (indiceCorrente == -1) ? 0 : indiceCorrente + 1;
        for (int i = partenza; i < listaDatiExcel.size(); i++) {
            if (!recordSalvatiManualmente.contains(i) && isRecordErrato(listaDatiExcel.get(i))) {
                indiceCorrente = i;
                popolaCampi(listaDatiExcel.get(i));
                aggiornaUI();
                return;
            }
        }
        JOptionPane.showMessageDialog(this, "Nessun altro errore trovato.");
    }

    private boolean isRecordErrato(Allegati a) {
        return checkError(a.getInizioLavori(), a) || checkError(a.getFineLavori(), a);
    }

    private boolean checkError(Date d, Allegati a) {
        if (d == null) return true;
        boolean isPI = a.getDescrizioneIntervento().toUpperCase().contains("PRONTO INTERVENTO");
        if (isPI) return !dateUguali(d, a.getDataOds());
        return isWeekend(d) || dateUguali(d, a.getDataOds()) || (a.getScadenzaOds() != null && d.after(a.getScadenzaOds()));
    }

    private void popolaCampi(Allegati a) {
        numeroOdsField.setText(a.getNumeroOds());
        dataOdsField.setText(a.getDataOds() != null ? sdf.format(a.getDataOds()) : "");
        scadenzaOdsField.setText(a.getScadenzaOds() != null ? sdf.format(a.getScadenzaOds()) : "");
        descrizioneInterventoField.setText(a.getDescrizioneIntervento());
        inizioLavoriField.setText(a.getInizioLavori() != null ? sdf.format(a.getInizioLavori()) : "");
        fineLavoriField.setText(a.getFineLavori() != null ? sdf.format(a.getFineLavori()) : "");

        checkFieldStyle(a.getInizioLavori(), inizioLavoriField, a, indiceCorrente);
        checkFieldStyle(a.getFineLavori(), fineLavoriField, a, indiceCorrente);
    }

    private void checkFieldStyle(Date d, JTextField field, Allegati a, int index) {
        if (d == null) { field.setBackground(java.awt.Color.WHITE); return; }
        if (recordSalvatiManualmente.contains(index)) {
            field.setBackground(new java.awt.Color(200, 255, 200));
            return;
        }
        boolean errore = checkError(d, a);
        field.setBackground(errore ? new java.awt.Color(255, 180, 180) : new java.awt.Color(210, 255, 210));
    }

    private void importaExcel() {
        JFileChooser fc = new JFileChooser(System.getProperty("user.home") + "/Desktop");
        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            try (FileInputStream fis = new FileInputStream(fc.getSelectedFile()); Workbook wb = new XSSFWorkbook(fis)) {
                Sheet sheet = wb.getSheetAt(0);
                listaDatiExcel.clear(); recordSalvatiManualmente.clear();
                DataFormatter df = new DataFormatter();
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row r = sheet.getRow(i); if (r == null) continue;
                    String ods = df.formatCellValue(r.getCell(2)).trim();
                    if (ods.isEmpty()) ods = "MANCANTE-" + (i + 1);
                    listaDatiExcel.add(new Allegati(ods, getCellValueAsDate(r.getCell(3)), getCellValueAsDate(r.getCell(4)),
                            df.formatCellValue(r.getCell(5)), df.formatCellValue(r.getCell(6)), df.formatCellValue(r.getCell(7)),
                            getCellValueAsDate(r.getCell(10)), getCellValueAsDate(r.getCell(11))));
                }
                indiceCorrente = 0; popolaCampi(listaDatiExcel.get(0)); aggiornaUI();
            } catch (Exception e) { e.printStackTrace(); }
        }
    }

    // --- UTILS ---
    private Date parseData(String str) throws Exception { if (str == null || str.trim().isEmpty()) return null; return sdf.parse(str.trim()); }
    private boolean isWeekend(Date d) { Calendar c = Calendar.getInstance(); c.setTime(d); int day = c.get(Calendar.DAY_OF_WEEK); return day == Calendar.SATURDAY || day == Calendar.SUNDAY; }
    private boolean dateUguali(Date d1, Date d2) { if (d1 == null || d2 == null) return false; return sdf.format(d1).equals(sdf.format(d2)); }
    private JTextField createStyledField() { JTextField tf = new JTextField(); tf.setFont(new java.awt.Font("Monospaced", java.awt.Font.PLAIN, 14)); return tf; }
    private void addLabelAndField(JPanel p, String l, JTextField f) { p.add(new JLabel(l)); p.add(f); }
    private void aggiornaUI() { precedenteButton.setEnabled(indiceCorrente > 0); prossimoButton.setEnabled(indiceCorrente < listaDatiExcel.size() - 1); infoExcelLabel.setText("Record: " + (indiceCorrente + 1) + " / " + listaDatiExcel.size()); }
    private void mostraProssimoDato() { if (indiceCorrente < listaDatiExcel.size() - 1) { indiceCorrente++; popolaCampi(listaDatiExcel.get(indiceCorrente)); aggiornaUI(); } }
    private void mostraDatoPrecedente() { if (indiceCorrente > 0) { indiceCorrente--; popolaCampi(listaDatiExcel.get(indiceCorrente)); aggiornaUI(); } }
    private void cercaOds() { String q = cercaOdsField.getText().trim(); for (int i = 0; i < listaDatiExcel.size(); i++) { if (listaDatiExcel.get(i).getNumeroOds().equalsIgnoreCase(q)) { indiceCorrente = i; popolaCampi(listaDatiExcel.get(i)); aggiornaUI(); return; } } }
    private Date getCellValueAsDate(Cell c) { if (c == null) return null; if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) return c.getDateCellValue(); try { return sdf.parse(c.getStringCellValue().trim()); } catch (Exception e) { return null; } }

    public static void main(String[] args) { SwingUtilities.invokeLater(() -> new Lavoro().setVisible(true)); }
}