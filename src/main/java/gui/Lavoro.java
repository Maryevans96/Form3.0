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
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Random;
import java.util.stream.Collectors;

public class Lavoro extends JFrame {

    private JPanel contentPane;
    private JTextField numeroOdsField, dataOdsField, scadenzaOdsField, viaField,
            danneggianteField, descrizioneInterventoField, inizioLavoriField, fineLavoriField;
    private JTextField cercaOdsField;

    private JButton scaricaButton, compilaButton, caricaExcelButton, prossimoButton, precedenteButton, pulisciCampiButton, cercaButton;
    private JCheckBox soloProntoInterventoCheckBox;
    private JLabel infoExcelLabel;

    private List<Allegati> listaDatiExcel = new ArrayList<>();
    private List<Allegati> listaAttuale = new ArrayList<>();
    private int indiceCorrente = -1;
    private String lastCompiledFilePath;
    private SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

    public Lavoro() {
        super("Compilatore PDF da Excel - Automazione ODS");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setPreferredSize(new Dimension(1100, 800));
        setMinimumSize(new Dimension(900, 700));

        contentPane = new JPanel(new BorderLayout(15, 15));
        contentPane.setBorder(new EmptyBorder(20, 20, 20, 20));

        // --- Pannello Superiore (Navigazione e Ricerca) ---
        JPanel topContainer = new JPanel(new GridLayout(2, 1, 5, 5));
        topContainer.setBorder(BorderFactory.createTitledBorder("Strumenti di Navigazione e Ricerca"));

        JPanel loadPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 5));
        caricaExcelButton = new JButton("Carica File Excel");
        precedenteButton = new JButton("<< Precedente");
        prossimoButton = new JButton("Prossimo >>");
        soloProntoInterventoCheckBox = new JCheckBox("Filtra Pronto Intervento");
        infoExcelLabel = new JLabel("Nessun file caricato");

        loadPanel.add(caricaExcelButton);
        loadPanel.add(precedenteButton);
        loadPanel.add(prossimoButton);
        loadPanel.add(soloProntoInterventoCheckBox);
        loadPanel.add(infoExcelLabel);

        JPanel searchPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 5));
        searchPanel.add(new JLabel("Cerca Numero O.d.S.:"));
        cercaOdsField = new JTextField(20);
        cercaButton = new JButton("Cerca e Vai");
        searchPanel.add(cercaOdsField);
        searchPanel.add(cercaButton);

        topContainer.add(loadPanel);
        topContainer.add(searchPanel);
        contentPane.add(topContainer, BorderLayout.NORTH);

        // --- Pannello Centrale (Form Dati) ---
        JPanel dataInputPanel = new JPanel(new GridLayout(0, 2, 10, 15));
        dataInputPanel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createTitledBorder("Anteprima Dati Estrazione"),
                new EmptyBorder(15, 15, 15, 15)
        ));

        java.awt.Font labelFont = new java.awt.Font("SansSerif", java.awt.Font.BOLD, 14);

        numeroOdsField = createStyledTextField();
        dataOdsField = createStyledTextField();
        scadenzaOdsField = createStyledTextField();
        viaField = createStyledTextField();
        danneggianteField = createStyledTextField();
        descrizioneInterventoField = createStyledTextField();
        inizioLavoriField = createStyledTextField();
        fineLavoriField = createStyledTextField();

        addLabelAndField(dataInputPanel, "Numero O.d.S.:", numeroOdsField, labelFont);
        addLabelAndField(dataInputPanel, "Data O.d.S.:", dataOdsField, labelFont);
        addLabelAndField(dataInputPanel, "Scadenza:", scadenzaOdsField, labelFont);
        addLabelAndField(dataInputPanel, "Via:", viaField, labelFont);
        addLabelAndField(dataInputPanel, "Danneggiante:", danneggianteField, labelFont);
        addLabelAndField(dataInputPanel, "Descrizione Intervento:", descrizioneInterventoField, labelFont);
        addLabelAndField(dataInputPanel, "Inizio Lavori:", inizioLavoriField, labelFont);
        addLabelAndField(dataInputPanel, "Fine Lavori:", fineLavoriField, labelFont);

        contentPane.add(dataInputPanel, BorderLayout.CENTER);

        // --- Pannello Inferiore (Azioni) ---
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 15, 10));
        compilaButton = new JButton("Genera PDF");
        scaricaButton = new JButton("Salva sul Desktop");
        pulisciCampiButton = new JButton("Pulisci Campi");
        JButton esciButton = new JButton("Esci");

        buttonPanel.add(compilaButton);
        buttonPanel.add(scaricaButton);
        buttonPanel.add(pulisciCampiButton);
        buttonPanel.add(esciButton);
        contentPane.add(buttonPanel, BorderLayout.SOUTH);

        // --- Event Listener ---
        caricaExcelButton.addActionListener(e -> importaExcel());
        prossimoButton.addActionListener(e -> mostraProssimoDato());
        precedenteButton.addActionListener(e -> mostraDatoPrecedente());
        soloProntoInterventoCheckBox.addActionListener(e -> applicaFiltro());
        cercaButton.addActionListener(e -> cercaOds());
        compilaButton.addActionListener(e -> compilePdf());
        scaricaButton.addActionListener(e -> downloadPdf());
        pulisciCampiButton.addActionListener(e -> clearFields());
        esciButton.addActionListener(e -> System.exit(0));

        add(contentPane);
        pack();
        setLocationRelativeTo(null);
    }

    private JTextField createStyledTextField() {
        JTextField tf = new JTextField();
        tf.setFont(new java.awt.Font("Monospaced", java.awt.Font.PLAIN, 15));
        return tf;
    }

    private void addLabelAndField(JPanel panel, String labelText, JTextField field, java.awt.Font font) {
        JLabel l = new JLabel(labelText);
        l.setFont(font);
        panel.add(l);
        panel.add(field);
    }

    private void importaExcel() {
        String desktopPath = System.getProperty("user.home") + File.separator + "Desktop";
        JFileChooser fileChooser = new JFileChooser(desktopPath);
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            try (FileInputStream fis = new FileInputStream(fileChooser.getSelectedFile());
                 Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(0);
                listaDatiExcel.clear();
                DataFormatter formatter = new DataFormatter();

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    String odsNum = formatter.formatCellValue(row.getCell(2)).trim();
                    if (odsNum.isEmpty()) continue;

                    String via = formatter.formatCellValue(row.getCell(5)).trim();
                    String dannegg = formatter.formatCellValue(row.getCell(6)).trim();
                    String descr = formatter.formatCellValue(row.getCell(7)).trim();
                    String note = formatter.formatCellValue(row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)).trim();
                    String descrizioneCompleta = descr + (note.isEmpty() ? "" : " - " + note);

                    Date dOds = getCellValueAsDate(row.getCell(3));
                    Date dScadenza = getCellValueAsDate(row.getCell(4));

                    // Recupero date esistenti (colonne 10 e 11)
                    Date dInizio = getCellValueAsDate(row.getCell(10));
                    Date dFine = getCellValueAsDate(row.getCell(11));

                    // LOGICA AUTOMATICA se le date sono vuote
                    if (dInizio == null || dFine == null) {
                        String check = descrizioneCompleta.toUpperCase();
                        if (check.contains("PRONTO INTERVENTO")) {
                            dInizio = dOds;
                            dFine = dOds;
                        } else if (check.contains("RIATTO ALLOGGIO")) {
                            Date[] dateRiatto = calcolaFinestraFeriale(dOds, dScadenza, 7);
                            if (dateRiatto != null) { dInizio = dateRiatto[0]; dFine = dateRiatto[1]; }
                        } else {
                            int durata = new Random().nextBoolean() ? 3 : 4;
                            Date[] dateStd = calcolaFinestraFeriale(dOds, dScadenza, durata);
                            if (dateStd != null) { dInizio = dateStd[0]; dFine = dateStd[1]; }
                        }
                    }

                    listaDatiExcel.add(new Allegati(odsNum, dOds, dScadenza, via, dannegg, descrizioneCompleta, dInizio, dFine));
                }

                if (!listaDatiExcel.isEmpty()) {
                    soloProntoInterventoCheckBox.setEnabled(true);
                    applicaFiltro();
                    JOptionPane.showMessageDialog(this, "Caricati " + listaDatiExcel.size() + " record.");
                }

            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Errore caricamento: " + ex.getMessage());
            }
        }
    }

    private Date[] calcolaFinestraFeriale(Date start, Date end, int durata) {
        if (start == null || end == null) return null;
        List<Date> feriali = new ArrayList<>();
        Calendar cal = Calendar.getInstance();
        cal.setTime(start);
        cal.add(Calendar.DAY_OF_MONTH, 1);

        while (cal.getTime().before(end)) {
            int dow = cal.get(Calendar.DAY_OF_WEEK);
            if (dow != Calendar.SATURDAY && dow != Calendar.SUNDAY) feriali.add(cal.getTime());
            cal.add(Calendar.DAY_OF_MONTH, 1);
        }
        if (feriali.isEmpty()) return null;
        if (feriali.size() < durata) durata = feriali.size();

        int startIndex = new Random().nextInt(feriali.size() - durata + 1);
        return new Date[]{ feriali.get(startIndex), feriali.get(startIndex + durata - 1) };
    }

    private void applicaFiltro() {
        if (soloProntoInterventoCheckBox.isSelected()) {
            listaAttuale = listaDatiExcel.stream()
                    .filter(a -> a.getDescrizioneIntervento().toUpperCase().contains("PRONTO INTERVENTO"))
                    .collect(Collectors.toList());
        } else {
            listaAttuale = new ArrayList<>(listaDatiExcel);
        }
        indiceCorrente = listaAttuale.isEmpty() ? -1 : 0;
        if (indiceCorrente != -1) popolaCampi(listaAttuale.get(0));
        aggiornaStatoBottoni();
    }

    private void aggiornaStatoBottoni() {
        precedenteButton.setEnabled(indiceCorrente > 0);
        prossimoButton.setEnabled(indiceCorrente < listaAttuale.size() - 1);
        infoExcelLabel.setText("Record: " + (indiceCorrente + 1) + " / " + listaAttuale.size());
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
        scaricaButton.setEnabled(false);
    }

    private void cercaOds() {
        String query = cercaOdsField.getText().trim();
        if (query.isEmpty() || listaAttuale.isEmpty()) return;
        for (int i = 0; i < listaAttuale.size(); i++) {
            if (listaAttuale.get(i).getNumeroOds().equalsIgnoreCase(query)) {
                indiceCorrente = i;
                popolaCampi(listaAttuale.get(i));
                aggiornaStatoBottoni();
                return;
            }
        }
        JOptionPane.showMessageDialog(this, "ODS non trovato.");
    }

    private void mostraProssimoDato() {
        if (indiceCorrente < listaAttuale.size() - 1) {
            indiceCorrente++;
            popolaCampi(listaAttuale.get(indiceCorrente));
            aggiornaStatoBottoni();
        }
    }

    private void mostraDatoPrecedente() {
        if (indiceCorrente > 0) {
            indiceCorrente--;
            popolaCampi(listaAttuale.get(indiceCorrente));
            aggiornaStatoBottoni();
        }
    }

    private Date getCellValueAsDate(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) return null;
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue();
        if (cell.getCellType() == CellType.STRING) {
            try { return sdf.parse(cell.getStringCellValue().trim()); } catch (ParseException e) { return null; }
        }
        return null;
    }

    private void clearFields() {
        numeroOdsField.setText(""); dataOdsField.setText(""); scadenzaOdsField.setText("");
        viaField.setText(""); danneggianteField.setText(""); descrizioneInterventoField.setText("");
        inizioLavoriField.setText(""); fineLavoriField.setText("");
        scaricaButton.setEnabled(false);
    }

    private void compilePdf() {
        try {
            Allegati dati = new Allegati(numeroOdsField.getText(),
                    (dataOdsField.getText().isEmpty() ? null : sdf.parse(dataOdsField.getText())),
                    (scadenzaOdsField.getText().isEmpty() ? null : sdf.parse(scadenzaOdsField.getText())),
                    viaField.getText(), danneggianteField.getText(), descrizioneInterventoField.getText(),
                    (inizioLavoriField.getText().isEmpty() ? null : sdf.parse(inizioLavoriField.getText())),
                    (fineLavoriField.getText().isEmpty() ? null : sdf.parse(fineLavoriField.getText())));

            File temp = File.createTempFile("preview_", ".pdf");
            temp.deleteOnExit();
            new PdfFiller().fillPdfSpecificFields(temp.getAbsolutePath(), dati);
            lastCompiledFilePath = temp.getAbsolutePath();
            scaricaButton.setEnabled(true);
            JOptionPane.showMessageDialog(this, "PDF generato correttamente!");
        } catch (Exception ex) { JOptionPane.showMessageDialog(this, "Errore: " + ex.getMessage()); }
    }

    private void downloadPdf() {
        JFileChooser saver = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");
        String safeName = numeroOdsField.getText().replaceAll("[\\\\/:*?\"<>|]", "_");
        saver.setSelectedFile(new File(safeName + ".pdf"));
        if (saver.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            try {
                java.nio.file.Files.copy(new File(lastCompiledFilePath).toPath(), saver.getSelectedFile().toPath(), java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                JOptionPane.showMessageDialog(this, "File salvato sul Desktop!");
            } catch (IOException e) { JOptionPane.showMessageDialog(this, "Errore salvataggio."); }
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new Lavoro().setVisible(true));
    }
}