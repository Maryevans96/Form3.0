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
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

public class Lavoro extends JFrame {

    private JPanel contentPane;
    private JTextField numeroOdsField, dataOdsField, scadenzaOdsField, viaField,
            danneggianteField, descrizioneInterventoField, inizioLavoriField, fineLavoriField;
    private JTextField cercaOdsField;

    private JButton caricaExcelButton, prossimoButton, precedenteButton,
            pulisciCampiButton, cercaButton, generaTuttiButton, esportaExcelButton;
    private JCheckBox soloProntoInterventoCheckBox;
    private JLabel infoExcelLabel;

    private List<Allegati> listaDatiExcel = new ArrayList<>();
    private List<Allegati> listaAttuale = new ArrayList<>();
    private int indiceCorrente = -1;
    private SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");

    public Lavoro() {
        super("Compilatore PDF Automazione 2026");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setPreferredSize(new Dimension(1100, 700));
        setMinimumSize(new Dimension(950, 600));

        contentPane = new JPanel(new BorderLayout(15, 15));
        contentPane.setBorder(new EmptyBorder(20, 20, 20, 20));

        // --- NORD: Strumenti ---
        JPanel topContainer = new JPanel(new GridLayout(2, 1, 5, 5));
        topContainer.setBorder(BorderFactory.createTitledBorder("Strumenti di Navigazione e Ricerca"));

        JPanel loadPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 5));
        caricaExcelButton = new JButton("Carica File Excel");
        esportaExcelButton = new JButton("Esporta Excel Aggiornato");
        esportaExcelButton.setBackground(new java.awt.Color(220, 220, 255));

        precedenteButton = new JButton("<< Precedente");
        prossimoButton = new JButton("Prossimo >>");
        soloProntoInterventoCheckBox = new JCheckBox("Filtra Pronto Intervento");
        infoExcelLabel = new JLabel("Nessun file caricato");

        loadPanel.add(caricaExcelButton);
        loadPanel.add(esportaExcelButton);
        loadPanel.add(precedenteButton);
        loadPanel.add(prossimoButton);
        loadPanel.add(soloProntoInterventoCheckBox);
        loadPanel.add(infoExcelLabel);

        JPanel searchPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 15, 5));
        searchPanel.add(new JLabel("Cerca O.d.S.:"));
        cercaOdsField = new JTextField(20);
        cercaButton = new JButton("Cerca");
        searchPanel.add(cercaOdsField);
        searchPanel.add(cercaButton);

        topContainer.add(loadPanel);
        topContainer.add(searchPanel);
        contentPane.add(topContainer, BorderLayout.NORTH);

        // --- CENTRO: Form ---
        JPanel dataInputPanel = new JPanel(new GridLayout(0, 2, 10, 15));
        dataInputPanel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createTitledBorder("Dati Intervento"),
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
        addLabelAndField(dataInputPanel, "Scadenza O.d.S.:", scadenzaOdsField, labelFont);
        addLabelAndField(dataInputPanel, "Via:", viaField, labelFont);
        addLabelAndField(dataInputPanel, "Danneggiante:", danneggianteField, labelFont);
        addLabelAndField(dataInputPanel, "Descrizione:", descrizioneInterventoField, labelFont);
        addLabelAndField(dataInputPanel, "Inizio Lavori:", inizioLavoriField, labelFont);
        addLabelAndField(dataInputPanel, "Fine Lavori:", fineLavoriField, labelFont);

        contentPane.add(dataInputPanel, BorderLayout.CENTER);

        // --- SUD: Azioni ---
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 15, 10));
        generaTuttiButton = new JButton("GENERA TUTTI PDF");
        generaTuttiButton.setFont(new java.awt.Font("SansSerif", java.awt.Font.BOLD, 14));
        generaTuttiButton.setBackground(new java.awt.Color(200, 230, 200));
        pulisciCampiButton = new JButton("Pulisci Campi");
        JButton esciButton = new JButton("Esci");

        // Disabilitati all'avvio
        precedenteButton.setEnabled(false);
        prossimoButton.setEnabled(false);
        generaTuttiButton.setEnabled(false);
        esportaExcelButton.setEnabled(false);
        soloProntoInterventoCheckBox.setEnabled(false);

        buttonPanel.add(generaTuttiButton);
        buttonPanel.add(pulisciCampiButton);
        buttonPanel.add(esciButton);
        contentPane.add(buttonPanel, BorderLayout.SOUTH);

        // Listeners
        caricaExcelButton.addActionListener(e -> importaExcel());
        esportaExcelButton.addActionListener(e -> esportaExcelAggiornato());
        prossimoButton.addActionListener(e -> mostraProssimoDato());
        precedenteButton.addActionListener(e -> mostraDatoPrecedente());
        soloProntoInterventoCheckBox.addActionListener(e -> applicaFiltro());
        cercaButton.addActionListener(e -> cercaOds());
        generaTuttiButton.addActionListener(e -> generaTuttiPdfMassivo());
        pulisciCampiButton.addActionListener(e -> clearFields());
        esciButton.addActionListener(e -> System.exit(0));

        add(contentPane);
        pack();
        setLocationRelativeTo(null);
    }

    private void importaExcel() {
        JFileChooser fileChooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            try (FileInputStream fis = new FileInputStream(fileChooser.getSelectedFile());
                 Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0);
                listaDatiExcel.clear();
                DataFormatter formatter = new DataFormatter();

                // Contatore per gestire lo slittamento ogni 5 ODS ordinari
                int contatoreOrdinari = 0;

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    if (row == null) continue;

                    String odsNumRaw = formatter.formatCellValue(row.getCell(2)).trim();
                    String via = formatter.formatCellValue(row.getCell(5)).trim();
                    String dannegg = formatter.formatCellValue(row.getCell(6)).trim();

                    if (odsNumRaw.isEmpty() && via.isEmpty() && dannegg.isEmpty()) continue;

                    String note = formatter.formatCellValue(row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)).trim();
                    String descBase = formatter.formatCellValue(row.getCell(7)).trim();
                    Date dOds = getCellValueAsDate(row.getCell(3));
                    Date dScadenza = getCellValueAsDate(row.getCell(4));

                    String odsNum = odsNumRaw.isEmpty() ? "RIGA-" + (i + 1) : odsNumRaw;

                    // --- LOGICA PRONTO INTERVENTO ---
                    boolean isPI = note.toUpperCase().contains("PRONTO INTERVENTO") ||
                            descBase.toUpperCase().contains("PRONTO INTERVENTO");

                    String descrizioneFinale = descBase + (!note.isEmpty() ? " - " + note : "");
                    if (isPI && !descrizioneFinale.toUpperCase().contains("PRONTO INTERVENTO")) {
                        descrizioneFinale += " - PRONTO INTERVENTO";
                    }

                    Date dInizio = null;
                    Date dFine = null;

                    if (dOds != null) {
                        Calendar calInizio = Calendar.getInstance();
                        calInizio.setTime(dOds);

                        if (isPI) {
                            // REGOLA PI: Inizio e Fine coincidono con ODS
                            dInizio = dOds;
                            dFine = dOds;
                        } else {
                            // 1. LOGICA SLITTAMENTO: Reset ogni 35 pratiche
                            int indiceCiclico = contatoreOrdinari % 35;
                            int giorniDaAggiungere = 1 + (indiceCiclico / 5);

                            // Calcolo Data Inizio
                            calInizio.add(Calendar.DAY_OF_MONTH, giorniDaAggiungere);
                            dInizio = calInizio.getTime();

                            // 2. Calcolo Data Fine (Default: Inizio + 2 giorni per un totale di 3)
                            Calendar calFine = Calendar.getInstance();
                            calFine.setTime(dInizio);
                            calFine.add(Calendar.DAY_OF_MONTH, 2);
                            dFine = calFine.getTime();

                            // 3. CONTROLLO DI SICUREZZA: Limite Scadenza ODS
                            if (dScadenza != null && dFine.after(dScadenza)) {
                                // Se la data calcolata sfora la scadenza, forziamo la fine alla scadenza
                                dFine = dScadenza;

                                // Ricalcoliamo l'inizio a ritroso dalla scadenza (3 giorni totali: Scadenza - 2)
                                Calendar calResetInizio = Calendar.getInstance();
                                calResetInizio.setTime(dFine);
                                calResetInizio.add(Calendar.DAY_OF_MONTH, -2);

                                // Se andando a ritroso finiamo prima della data dell'ODS, blocchiamo su ODS
                                if (calResetInizio.getTime().before(dOds)) {
                                    dInizio = dOds;
                                } else {
                                    dInizio = calResetInizio.getTime();
                                }
                            }

                            // Incrementiamo solo per i non-PI
                            contatoreOrdinari++;
                        }
                    }

                    listaDatiExcel.add(new Allegati(odsNum, dOds, dScadenza, via, dannegg, descrizioneFinale, dInizio, dFine));
                }

                if (!listaDatiExcel.isEmpty()) {
                    soloProntoInterventoCheckBox.setEnabled(true);
                    generaTuttiButton.setEnabled(true);
                    esportaExcelButton.setEnabled(true);
                    applicaFiltro();
                    JOptionPane.showMessageDialog(this, "Caricate e elaborate " + listaDatiExcel.size() + " righe.");
                }
            } catch (Exception ex) {
                ex.printStackTrace(); // Utile per il debug in console
                JOptionPane.showMessageDialog(this, "Errore caricamento: " + ex.getMessage());
            }
        }
    }

    private void esportaExcelAggiornato() {
        if (listaDatiExcel.isEmpty()) return;
        JFileChooser saver = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");
        saver.setSelectedFile(new File("Report_Lavori_Compilati_2026.xlsx"));

        if (saver.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            try (Workbook workbook = new XSSFWorkbook();
                 FileOutputStream fos = new FileOutputStream(saver.getSelectedFile())) {

                Sheet sheet = workbook.createSheet("Dati Elaborati");
                String[] headers = {"Num ODS", "Data ODS", "Scadenza", "Via", "Danneggiante", "Descrizione", "Inizio Lavori", "Fine Lavori"};

                Row headerRow = sheet.createRow(0);
                CellStyle headerStyle = workbook.createCellStyle();
                org.apache.poi.ss.usermodel.Font poiFont = workbook.createFont();
                poiFont.setBold(true);
                headerStyle.setFont(poiFont);

                for (int i = 0; i < headers.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(headers[i]); cell.setCellStyle(headerStyle);
                }

                CellStyle dateStyle = workbook.createCellStyle();
                dateStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd/mm/yyyy"));

                int rowNum = 1;
                for (Allegati a : listaDatiExcel) {
                    Row row = sheet.createRow(rowNum++);
                    row.createCell(0).setCellValue(a.getNumeroOds());
                    addDateCell(row, 1, a.getDataOds(), dateStyle);
                    addDateCell(row, 2, a.getScadenzaOds(), dateStyle);
                    row.createCell(3).setCellValue(a.getVia());
                    row.createCell(4).setCellValue(a.getDanneggiante());
                    row.createCell(5).setCellValue(a.getDescrizioneIntervento());
                    addDateCell(row, 6, a.getInizioLavori(), dateStyle);
                    addDateCell(row, 7, a.getFineLavori(), dateStyle);
                }

                for (int i = 0; i < headers.length; i++) sheet.autoSizeColumn(i);
                workbook.write(fos);
                JOptionPane.showMessageDialog(this, "File Excel esportato con successo sul Desktop!");
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Errore durante l'esportazione: " + ex.getMessage());
            }
        }
    }

    private void addDateCell(Row row, int col, Date date, CellStyle style) {
        if (date != null) {
            Cell cell = row.createCell(col);
            cell.setCellValue(date);
            cell.setCellStyle(style);
        }
    }

    private void generaTuttiPdfMassivo() {
        if (listaAttuale.isEmpty()) return;
        JFileChooser chooser = new JFileChooser(System.getProperty("user.home") + File.separator + "Desktop");
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File folder = chooser.getSelectedFile();
            PdfFiller filler = new PdfFiller();
            int ok = 0, err = 0;
            for (Allegati a : listaAttuale) {
                try {
                    String name = a.getNumeroOds().replaceAll("[\\\\/:*?\"<>|]", "_") + ".pdf";
                    filler.fillPdfSpecificFields(new File(folder, name).getAbsolutePath(), a);
                    ok++;
                } catch (Exception ex) { err++; }
            }
            JOptionPane.showMessageDialog(this, "Generazione completata.\nPDF creati: " + ok + "\nErrori: " + err);
        }
    }

    private JTextField createStyledTextField() {
        JTextField tf = new JTextField();
        tf.setFont(new java.awt.Font("Monospaced", java.awt.Font.PLAIN, 15));
        return tf;
    }

    private void addLabelAndField(JPanel panel, String label, JTextField field, java.awt.Font font) {
        JLabel l = new JLabel(label); l.setFont(font);
        panel.add(l); panel.add(field);
    }

    private void cercaOds() {
        String q = cercaOdsField.getText().trim();
        if (q.isEmpty()) return;
        for (int i = 0; i < listaAttuale.size(); i++) {
            if (listaAttuale.get(i).getNumeroOds().equalsIgnoreCase(q)) {
                indiceCorrente = i; popolaCampi(listaAttuale.get(i)); aggiornaStatoBottoni();
                return;
            }
        }
        JOptionPane.showMessageDialog(this, "O.d.S. non trovato nella lista attuale.");
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
        else clearFields();
        aggiornaStatoBottoni();
    }

    private void aggiornaStatoBottoni() {
        precedenteButton.setEnabled(indiceCorrente > 0);
        prossimoButton.setEnabled(indiceCorrente < listaAttuale.size() - 1);
        infoExcelLabel.setText("Record: " + (indiceCorrente + 1) + " / " + listaAttuale.size());
    }

    private void mostraProssimoDato() {
        if (indiceCorrente < listaAttuale.size() - 1) {
            indiceCorrente++; popolaCampi(listaAttuale.get(indiceCorrente)); aggiornaStatoBottoni();
        }
    }

    private void mostraDatoPrecedente() {
        if (indiceCorrente > 0) {
            indiceCorrente--; popolaCampi(listaAttuale.get(indiceCorrente)); aggiornaStatoBottoni();
        }
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

    private Date getCellValueAsDate(Cell cell) {
        if (cell == null || cell.getCellType() == CellType.BLANK) return null;
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) return cell.getDateCellValue();
        if (cell.getCellType() == CellType.STRING) {
            try { return sdf.parse(cell.getStringCellValue().trim()); } catch (ParseException e) { }
        }
        return null;
    }

    private void clearFields() {
        numeroOdsField.setText(""); dataOdsField.setText(""); scadenzaOdsField.setText("");
        viaField.setText(""); danneggianteField.setText(""); descrizioneInterventoField.setText("");
        inizioLavoriField.setText(""); fineLavoriField.setText("");
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new Lavoro().setVisible(true));
    }
}