package org.example;

import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.convert.out.pdf.PdfConversion;
import org.docx4j.convert.out.pdf.viaXSLFO.Conversion;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main extends JFrame {
    private JTextField outputDirField;
    private JFileChooser fileChooser;
    private DefaultListModel<File> fileListModel;

    public Main() {
        setTitle("Word to PDF Converter");
        setSize(600, 400);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLocationRelativeTo(null);

        // Initialize file chooser and set file filter for Word files
        fileChooser = new JFileChooser();
        fileChooser.setMultiSelectionEnabled(true);
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        FileFilter wordFileFilter = new FileNameExtensionFilter("Word Documents", "doc", "docx");
        fileChooser.setFileFilter(wordFileFilter);

        fileListModel = new DefaultListModel<>();
        JList<File> fileList = new JList<>(fileListModel);
        JScrollPane fileListScrollPane = new JScrollPane(fileList);

        JButton addButton = new JButton("Add Files");
        addButton.setPreferredSize(new Dimension(addButton.getPreferredSize().width, 50));
        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                if (fileChooser.showOpenDialog(Main.this) == JFileChooser.APPROVE_OPTION) {
                    for (File file : fileChooser.getSelectedFiles()) {
                        fileListModel.addElement(file);
                    }
                }
            }
        });

        JButton removeButton = new JButton("Remove Selected");
        removeButton.setPreferredSize(new Dimension(removeButton.getPreferredSize().width, 50));
        removeButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                fileListModel.removeElement(fileList.getSelectedValue());
            }
        });

        JButton convertButton = new JButton("Convert to PDF");
        convertButton.setPreferredSize(new Dimension(convertButton.getPreferredSize().width, 50));
        convertButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                convertFilesToPdf();
            }
        });

        JLabel outputDirLabel = new JLabel("Output Directory:");
        outputDirField = new JTextField();
        JButton outputDirButton = new JButton("Browse");
        outputDirButton.setPreferredSize(new Dimension(outputDirButton.getPreferredSize().width, 30));
        outputDirButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                chooseOutputDirectory();
            }
        });

        // Padding and spacing
        fileListScrollPane.setBorder(new EmptyBorder(0, 10, 0, 10));

        JPanel buttonPanel = new JPanel();
        buttonPanel.setBorder(new EmptyBorder(5, 10, 5, 10));
        buttonPanel.add(addButton);
        buttonPanel.add(removeButton);
        buttonPanel.add(convertButton);

        JPanel outputDirPanel = new JPanel(new BorderLayout(10, 10)); // 10px horizontal and vertical gaps
        outputDirPanel.setBorder(new EmptyBorder(10, 10, 10, 10));
        outputDirPanel.add(outputDirLabel, BorderLayout.WEST);
        outputDirPanel.add(outputDirField, BorderLayout.CENTER);
        outputDirPanel.add(outputDirButton, BorderLayout.EAST);

        JPanel panel = new JPanel(new BorderLayout(10, 10)); // 10px horizontal and vertical gaps
        panel.setBorder(new EmptyBorder(10, 10, 10, 10));
        panel.add(fileListScrollPane, BorderLayout.CENTER);
        panel.add(buttonPanel, BorderLayout.SOUTH);
        panel.add(outputDirPanel, BorderLayout.NORTH);

        getContentPane().add(panel, BorderLayout.CENTER);
    }

    private void chooseOutputDirectory() {
        JFileChooser directoryChooser = new JFileChooser();
        directoryChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        if (directoryChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            outputDirField.setText(directoryChooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void convertFilesToPdf() {
        String outputDir = outputDirField.getText();
        if (outputDir.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please specify an output directory.", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }

        for (int i = 0; i < fileListModel.size(); i++) {
            File docxFile = fileListModel.getElementAt(i);
            File pdfFile = new File(outputDir, docxFile.getName().replace(".docx", ".pdf").replace(".doc", ".pdf"));
            try {
                WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(docxFile);

                // Map fonts
                mapFonts(wordMLPackage);

                // Convert to PDF
                PdfConversion conversion = new Conversion(wordMLPackage);
                PdfSettings pdfSettings = new PdfSettings();
                conversion.output(new FileOutputStream(pdfFile), pdfSettings);

                JOptionPane.showMessageDialog(this, "File converted successfully: " + pdfFile.getName(), "Success", JOptionPane.INFORMATION_MESSAGE);

            } catch (Docx4JException | IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this, "Error converting file: " + docxFile.getName(), "Error", JOptionPane.ERROR_MESSAGE);
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }

    private void mapFonts(WordprocessingMLPackage wordMLPackage) throws Exception {
        Mapper fontMapper = new IdentityPlusMapper();

        // Adding a specific font mapping
        PhysicalFont timesNewRoman = PhysicalFonts.get("Times New Roman");
        if (timesNewRoman != null) {
            fontMapper.put("Times New Roman", timesNewRoman);
        }

        // Example of adding more fonts
        PhysicalFont arial = PhysicalFonts.get("Arial");
        if (arial != null) {
            fontMapper.put("Arial", arial);
        }

        // Set the font mapper to the WordprocessingMLPackage
        wordMLPackage.setFontMapper(fontMapper);
    }

    public static void main(String[] args) {
        try {
            // Set the FlatLaf look and feel
            UIManager.setLookAndFeel(new com.formdev.flatlaf.FlatLightLaf());
        } catch (UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }

        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new Main().setVisible(true);
            }
        });
    }
}
