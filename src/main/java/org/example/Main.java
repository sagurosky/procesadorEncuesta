package org.example;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.text.DefaultCaret;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.Date;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Main {
    private StringBuilder logHtml = new StringBuilder("<html><body style='color:white; font-family:Consolas; font-size:13px;'>");

    private JFrame frame;
    private JTextPane logArea;
    private JProgressBar progressBar;
    private File archivoEncuesta;
    private File archivoTemplate;
    private JCheckBox checkboxCSV;
    private JCheckBox checkboxNoDescargarImagenes;
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new Main().crearInterfaz());
    }

    // 🔹 Fragmento modificado de tu clase Main.java
// (solo cambia la parte de crearInterfaz y se agrega el método limpiarLogs())

    private void crearInterfaz() {
        frame = new JFrame("📊 Procesador de Encuestas");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(700, 500);
        frame.setLayout(new BorderLayout(10, 10));
        frame.getContentPane().setBackground(new Color(240, 240, 240));

        JPanel panelBotones = new JPanel(new GridLayout(1, 3, 10, 10));
        panelBotones.setBorder(new EmptyBorder(10, 10, 10, 10));
        panelBotones.setBackground(new Color(240, 240, 240));

        // 🔹 Panel para checkboxes y botón limpiar
        JPanel panelOpciones = new JPanel();
        panelOpciones.setLayout(new BoxLayout(panelOpciones, BoxLayout.X_AXIS));
        panelOpciones.setBorder(new EmptyBorder(0, 10, 10, 10));
        panelOpciones.setBackground(new Color(240, 240, 240));

        // Checkboxes
        checkboxCSV = new JCheckBox("Exportar como CSV");
        checkboxNoDescargarImagenes = new JCheckBox("No descargar imágenes");
        checkboxCSV.setBackground(new Color(240, 240, 240));
        checkboxNoDescargarImagenes.setBackground(new Color(240, 240, 240));

        // 🔹 Botón limpiar logs
        JButton btnLimpiarLogs = new JButton("🧹 Limpiar logs");
        btnLimpiarLogs.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        btnLimpiarLogs.setBackground(new Color(200, 200, 200));
        btnLimpiarLogs.setFocusPainted(false);
        btnLimpiarLogs.setBorder(BorderFactory.createEmptyBorder(5, 10, 5, 10));
        btnLimpiarLogs.addActionListener(e -> limpiarLogs());

        // Agregar checkboxes y botón al panel horizontal
        panelOpciones.add(Box.createHorizontalGlue());
        panelOpciones.add(checkboxCSV);
        panelOpciones.add(Box.createHorizontalStrut(15));
        panelOpciones.add(checkboxNoDescargarImagenes);
        panelOpciones.add(Box.createHorizontalStrut(20));
        panelOpciones.add(btnLimpiarLogs);
        panelOpciones.add(Box.createHorizontalGlue());

        // Panel superior (botones arriba, opciones abajo)
        JPanel panelSuperior = new JPanel(new BorderLayout());
        panelSuperior.setBackground(new Color(240, 240, 240));
        panelSuperior.add(panelBotones, BorderLayout.NORTH);
        panelSuperior.add(panelOpciones, BorderLayout.SOUTH);

        JButton btnEncuesta = crearBoton("📥 Encuesta (.xlsx)", new Color(70, 130, 180));
        JButton btnTemplate = crearBoton("📄 Template (.xls)", new Color(60, 179, 113));
        JButton btnProcesar = crearBoton("🚀 Procesar", new Color(255, 140, 0));

        panelBotones.add(btnEncuesta);
        panelBotones.add(btnTemplate);
        panelBotones.add(btnProcesar);

        progressBar = new JProgressBar(0, 100);
        progressBar.setStringPainted(true);
        progressBar.setPreferredSize(new Dimension(400, 30));
        progressBar.setForeground(new Color(0, 153, 76));
        progressBar.setBackground(new Color(230, 230, 230));
        progressBar.setValue(0);

        logArea = new JTextPane();
        logArea.setContentType("text/html");
        logArea.setEditable(false);
        logArea.setBorder(BorderFactory.createLineBorder(Color.GRAY));
        logArea.setBackground(Color.BLACK);
        logArea.setForeground(Color.WHITE);
        logArea.setFont(new Font("Consolas", Font.PLAIN, 14));

        JScrollPane scrollLog = new JScrollPane(logArea);
        scrollLog.setBorder(new EmptyBorder(10, 10, 10, 10));

        logArea.setText("<html><body style='color:white; font-family:Consolas; font-size:13px; padding:10px;'>"
                + "<b>📝 Log de proceso:</b><br></body></html>");

        DefaultCaret caret = (DefaultCaret) logArea.getCaret();
        caret.setUpdatePolicy(DefaultCaret.ALWAYS_UPDATE);

        frame.add(panelSuperior, BorderLayout.NORTH);
        frame.add(scrollLog, BorderLayout.CENTER);
        frame.add(progressBar, BorderLayout.SOUTH);

        btnEncuesta.addActionListener(e -> seleccionarArchivo(true));
        btnTemplate.addActionListener(e -> seleccionarArchivo(false));
        btnProcesar.addActionListener(e -> procesar());

        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }

    // 🔹 Nuevo método para limpiar logs
    private void limpiarLogs() {
        logHtml = new StringBuilder("<html><body style='color:white; font-family:Consolas; font-size:13px;'>");
        logHtml.append("<b>📝 Log de proceso:</b><br>");
        logArea.setText(logHtml.toString() + "</body></html>");
//        logInfo("🧹 Logs limpiados correctamente.");
    }


    private JButton crearBoton(String texto, Color color) {
        JButton boton = new JButton(texto);
        boton.setFont(new Font("Segoe UI", Font.BOLD, 16));
        boton.setBackground(color);
        boton.setForeground(Color.WHITE);
        boton.setFocusPainted(false);
        boton.setPreferredSize(new Dimension(180, 50));
        return boton;
    }

    private void seleccionarArchivo(boolean esEncuesta) {
//        JFileChooser chooser = new JFileChooser();
        JFileChooser chooser = new JFileChooser(new File("."));

        int resultado = chooser.showOpenDialog(frame);
        if (resultado == JFileChooser.APPROVE_OPTION) {
            File archivo = chooser.getSelectedFile();
            String nombre = archivo.getName().toLowerCase();

            if (esEncuesta && !nombre.endsWith(".xlsx")) {
                logError("El archivo de encuesta debe ser .xlsx");
                return;
            }
            if (!esEncuesta && !nombre.endsWith(".xls")) {
                logError("El archivo template debe ser .xls (versión 97-2003)");
                return;
            }

            if (esEncuesta) {
                archivoEncuesta = archivo;
                logInfo("📥 Archivo encuesta seleccionado: " + archivo.getAbsolutePath());
            } else {
                archivoTemplate = archivo;
                logInfo("📄 Archivo template seleccionado: " + archivo.getAbsolutePath());
            }
        }
    }

    private void procesar() {
        if (archivoEncuesta == null || archivoTemplate == null) {
            logError("Debe seleccionar ambos archivos antes de procesar.");
            return;
        }

        logInfo("🚀 Procesando...");
        progressBar.setVisible(true);
        progressBar.setIndeterminate(true);

        SwingWorker<Void, Void> worker = new SwingWorker<>() {
            @Override
            protected Void doInBackground() {
                try {
                    File carpetaProceso = new File("proceso/imagenes");
                    if (!carpetaProceso.exists()) {
                        boolean creado = carpetaProceso.mkdirs();
                        if (!creado) {
                            logError("No se pudo crear la carpeta 'proceso/imagenes'.");
                            return null;
                        }
                    }

                    logInfo("✔️ Carpeta de procesamiento lista.");

                    if (!validarPreguntas()) {
                        logError("❌ Las preguntas no coinciden. Proceso detenido.");
                        return null;
                    }

                    logInfo("✔️ Validación exitosa. Generando archivos...");
                    procesarArchivos();



                } catch (Exception ex) {
                    logError("Error: " + ex.getMessage());
                }
                return null;
            }

            @Override
            protected void done() {
                progressBar.setVisible(false);
                progressBar.setIndeterminate(false);
                logInfo("✅ Proceso finalizado.");
            }
        };

        worker.execute();
    }

    private boolean validarPreguntas() {
        try (FileInputStream fisEncuesta = new FileInputStream(archivoEncuesta);
             FileInputStream fisTemplate = new FileInputStream(archivoTemplate)) {

            XSSFWorkbook wbEncuesta = new XSSFWorkbook(fisEncuesta);
            HSSFWorkbook wbTemplate = new HSSFWorkbook(fisTemplate);

            Sheet hojaEncuesta = wbEncuesta.getSheetAt(0);
            Sheet hojaTemplate = wbTemplate.getSheetAt(0);



            Row filaEncuesta = hojaEncuesta.getRow(0); // Fila 2 (índice 1)
            if (filaEncuesta == null) {
                logError("La fila 2 del archivo encuesta está vacía.");
                return false;
            }
            int cantidadPreguntas = obtenerCantidadPreguntasValidas(filaEncuesta, hojaTemplate);
            if(cantidadPreguntas==-1)return false;
            for (int i = 0; i < cantidadPreguntas; i++) { //  preguntas esperadas
                Cell celdaEncuesta = filaEncuesta.getCell(7 + i); // desde H (7) hasta AD (29)
                Row filaTemplate = hojaTemplate.getRow(1 + i);    //
                if (filaTemplate == null) {
                    logError("La fila " + (2 + i) + " del template está vacía.");
                    return false;
                }
                Cell celdaTemplate = filaTemplate.getCell(1); // columna B = índice 1

                String valorEncuesta = celdaEncuesta != null ? celdaEncuesta.toString().trim() : "";
                String valorTemplate = celdaTemplate != null ? celdaTemplate.toString().trim() : "";

                if (!valorEncuesta.equalsIgnoreCase(valorTemplate)) {
                    logError("Diferencia en la pregunta " + (i + 1) + ":");
                    logError("→ Encuesta (columna " + getColumnaExcel(7 + i) + "2): '" + valorEncuesta + "'");
                    logError("→ Template (celda B" + (2 + i) + "): '" + valorTemplate + "'");
                    return false;
                }
            }

            wbEncuesta.close();
            wbTemplate.close();

            return true;

        } catch (Exception ex) {
            logError("Error al validar preguntas: " + ex.getMessage());
            ex.printStackTrace();
            return false;
        }
    }

    private String getColumnaExcel(int index) {
        StringBuilder col = new StringBuilder();
        while (index >= 0) {
            col.insert(0, (char) ('A' + (index % 26)));
            index = index / 26 - 1;
        }
        return col.toString();
    }


    private void logInfo(String mensaje) {
        agregarAlLog("<span style='color:white;'>" + mensaje + "</span>");
    }


    private void logError(String mensaje) {
        agregarAlLog("<span style='color:red;'>[ERROR]</span> " + mensaje);
    }

    private void agregarAlLog(String mensaje) {
        logHtml.append(mensaje).append("<br>");
        logArea.setText(logHtml.toString() + "</body></html>");
    }
    private void procesarArchivos() {


        if (archivoEncuesta == null || archivoTemplate == null) {
            logError("Debe seleccionar ambos archivos antes de procesar.");
            return;
        }

        try (FileInputStream fisEncuesta = new FileInputStream(archivoEncuesta);
             FileInputStream fisTemplate = new FileInputStream(archivoTemplate)) {

            XSSFWorkbook wbEncuesta = new XSSFWorkbook(fisEncuesta);
            HSSFWorkbook wbTemplateOriginal = new HSSFWorkbook(fisTemplate);

            Sheet sheetEncuesta = wbEncuesta.getSheetAt(0);
            Sheet sheetPreguntasTemplate = wbTemplateOriginal.getSheetAt(0);

            // Validación de preguntas
            Row filaEncuesta = sheetEncuesta.getRow(0); // G1 - AD1 (índice 0)
            if (filaEncuesta == null) {
                logError("No se pudo acceder a la fila de preguntas del archivo encuesta.");
                return;
            }
            int cantidadPreguntas = obtenerCantidadPreguntasValidas(filaEncuesta, sheetPreguntasTemplate);
            if (cantidadPreguntas == -1) return; // ya logueamos el error

            for (int i = 0; i < cantidadPreguntas; i++) {
                Row filaTemplate = sheetPreguntasTemplate.getRow(i + 1); // Fila 2 a 25
                if (filaTemplate == null) {
                    logError("No se pudo acceder a la fila de preguntas del template en la posición " + (i + 2));
                    return;
                }

                Cell celdaEncuesta = filaEncuesta.getCell(i + 7); // Columnas H (7) a AD (29)
                Cell celdaTemplate = filaTemplate.getCell(1);     // Columna B (1)

                String valorEncuesta = obtenerValorCeldaComoTexto(celdaEncuesta).trim();
                String valorTemplate = obtenerValorCeldaComoTexto(celdaTemplate).trim();

                if (!valorEncuesta.equals(valorTemplate)) {
                    logError("Diferencia de pregunta en posición " + (i + 1) + ": '" +
                            valorEncuesta + "' ≠ '" + valorTemplate + "'");
                    return;
                }
            }

            // Crear carpetas si no existen
            File carpetaProceso = new File("proceso");
            if (!carpetaProceso.exists()) carpetaProceso.mkdir();

            File carpetaImagenes = new File(carpetaProceso, "imagenes");
            if (!carpetaImagenes.exists()) carpetaImagenes.mkdir();

            int totalFilas = sheetEncuesta.getLastRowNum();
            int contadorImagen = 1;

            SwingUtilities.invokeLater(() -> {
                progressBar.setMinimum(0);
                progressBar.setMaximum(totalFilas - 1);
                progressBar.setValue(0);
            });

            for (int filaIndex = 1; filaIndex <= totalFilas; filaIndex++) {
                Row fila = sheetEncuesta.getRow(filaIndex);
//                if (fila == null) continue;
                if (filaEstaVacia(fila)) continue;

                // Validar columna "Revisado" (columna E -> índice 4)
                String valorRevisado = obtenerValorCeldaComoTexto(fila.getCell(4)).trim().toLowerCase();
                if (!(valorRevisado.equals("si") || valorRevisado.equals("sí"))) {
                    logInfo("Fila " + (filaIndex + 1) + ": sin revisar (valor en columna \"Revisado\" distinto de 'sí')");
                    continue;
                }


                // Obtener nombre de archivo desde columna F (índice 5)
                String nombreArchivo = obtenerValorCeldaComoTexto(fila.getCell(5)).trim();
                if (nombreArchivo.isEmpty()) {
                    logError("Fila " + (filaIndex + 1) + ": nombre de archivo vacío.");
                    continue;
                }

                // Copiar workbook template original
                HSSFWorkbook nuevoTemplate = new HSSFWorkbook();
                Sheet hojaNueva = nuevoTemplate.createSheet("Hoja1");
                copiarContenidoHoja(wbTemplateOriginal.getSheetAt(0), hojaNueva);

                for (int i = 0; i < cantidadPreguntas; i++) {
                    Row filaEncuestaFoto = sheetEncuesta.getRow(filaIndex);
                    Cell celdaEncuesta = filaEncuestaFoto.getCell(7 + i); // desde H (7) hasta AD (29)
                    Row filaTemplate = hojaNueva.getRow(i + 1); // Fila 2 a 25
                    if (filaTemplate == null) filaTemplate = hojaNueva.createRow(i + 1);

                    Cell celdaPregunta = fila.getCell(i + 7); // H2 (7) a AD2 (29)
                    String valor = obtenerValorCeldaComoTexto(celdaPregunta).trim();

                    Cell tipoCelda = filaTemplate.getCell(0); // Col A = tipo
                    String tipo = obtenerValorCeldaComoTexto(tipoCelda).trim().toUpperCase();

                    Cell celdaDestino = filaTemplate.getCell(2); // Col C
                    if (celdaDestino == null) celdaDestino = filaTemplate.createCell(2);
                    Cell celdaVacia = filaTemplate.getCell(3); // Col D
                    if (celdaVacia == null) celdaVacia = filaTemplate.createCell(3);

                    boolean estaVacio = valor.isEmpty();
                    celdaVacia.setCellValue(estaVacio ? "True" : "False");

                    switch (tipo) {
                        case "TEXTO":
                        case "LISTA":
                            celdaDestino.setCellValue(valor.isEmpty()?"Sin respuesta":valor);
                            break;

                        case "BINOMIAL":
                            String normalizado = valor.trim().toLowerCase(Locale.ROOT).replace("í", "i");
                            if (normalizado.equals("si") || normalizado.equals("no")) {
                                celdaDestino.setCellValue(Character.toUpperCase(normalizado.charAt(0)) + normalizado.substring(1));
                            } else {
                                logError("Fila " + (filaIndex + 1) + ", pregunta " + (i + 1) + ": valor inválido BINOMIAL: " + valor);
                                return;
                            }
                            break;

                        case "NUMERICA":
                            try {
                                double num = Double.parseDouble(valor);
                                celdaDestino.setCellValue(num);
                                celdaVacia.setCellValue("False");
                            } catch (NumberFormatException e) {
                                celdaDestino.setCellValue(valor.isEmpty()?"Sin respuesta":valor);
                                celdaVacia.setCellValue("True");
                            }
                            break;

                        case "FOTO":





                            String urlOriginal = obtenerValorCeldaComoTexto(celdaEncuesta).trim(); // celdaEncuesta es la celda actual del archivo encuesta
                            if (urlOriginal.isEmpty()) {
                                celdaVacia.setCellValue("True");  // No hay URL => true (falta de respuesta)
                                celdaDestino.setCellValue("Sin respuesta");
                                break;
                            }

                            // Hay URL, intentamos descargar
                            try {
                                Pattern pattern = Pattern.compile("(?<=id=|/d/)([a-zA-Z0-9_-]{10,})");
                                Matcher matcher = pattern.matcher(urlOriginal);
                                if (!matcher.find()) {
                                    logError("Fila " + (filaIndex + 1) + ": enlace inválido de Google Drive: " + urlOriginal);
                                    celdaVacia.setCellValue("True"); // Lo tratamos como fallido
                                    break;
                                }
                                String fileId = matcher.group(1);
                                String urlDescarga = "https://drive.google.com/uc?export=download&id=" + fileId;
                                String nombreImagen="";
                                if (!checkboxNoDescargarImagenes.isSelected()) {
                                    BufferedImage imagen = ImageIO.read(new URL(urlDescarga));
                                    if (imagen == null) {
                                        String sucursal = obtenerValorCeldaComoTexto(sheetEncuesta.getRow(filaIndex).getCell(5)).trim(); // Columna F
                                        String columnaLetra = convertirIndiceAColumnaExcel(7 + i); // Convierte índice a letra (ej. 7 = H)
                                        logError("Fila " + (filaIndex + 1) + ", columna " + columnaLetra + " (Sucursal: " + sucursal + "): no se pudo obtener una imagen válida desde el enlace.");
                                        celdaVacia.setCellValue("True");
                                        break;
                                    }

                                     nombreImagen = nombreArchivo + "_" + LocalDate.now() + "_R" + (contadorImagen++) + ".jpg";
                                    File archivoImagen = new File(carpetaImagenes, nombreImagen);
                                    ImageIO.write(imagen, "jpg", archivoImagen);
                                }else{
                                     nombreImagen = nombreArchivo + "_" + LocalDate.now() + "_R" + (contadorImagen++) + ".jpg";
                                }

                                celdaDestino.setCellValue(nombreImagen);
                                celdaVacia.setCellValue("False"); // Se descargó correctamente
                            } catch (Exception ex) {
                                logError("Fila " + (filaIndex + 1) + ": error al descargar imagen. " + ex.getMessage());
                                celdaVacia.setCellValue("True");
                            }
                            break;



                        default:
                            logError("Tipo de pregunta desconocido en fila template " + (i + 2) + ": " + tipo);
                            return;
                    }
                }

                if (checkboxCSV.isSelected()) {
                    File archivoSalidaCSV = new File(carpetaProceso, nombreArchivo + ".csv");
                    try (BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(archivoSalidaCSV), StandardCharsets.UTF_8))) {
                        writer.write("Tipo;Pregunta;Respuesta;Saltea");
                        writer.newLine();

                        for (int i = 0; i < cantidadPreguntas; i++) {
                            Row filaHoja = hojaNueva.getRow(i + 1);
                            if (filaHoja == null) continue;

                            String tipo = limpiarSaltosDeLinea(obtenerValorCeldaComoTexto(filaHoja.getCell(0)));
                            String pregunta = limpiarSaltosDeLinea(obtenerValorCeldaComoTexto(filaHoja.getCell(1)));
                            String respuesta = limpiarSaltosDeLinea(obtenerValorCeldaComoTexto(filaHoja.getCell(2)));
                            String vacio = limpiarSaltosDeLinea(obtenerValorCeldaComoTexto(filaHoja.getCell(3)));

                            // CSV con 4 columnas por filaHoja: tipo, pregunta, respuesta, está vacío
                            writer.write(String.format("%s;%s;%s;%s", tipo, pregunta, respuesta, vacio));

                            writer.newLine();
                        }
                    }
                } else {
                    File archivoSalida = new File(carpetaProceso, nombreArchivo + ".xls");
                    try (FileOutputStream fos = new FileOutputStream(archivoSalida)) {
                        nuevoTemplate.write(fos);
                    }
                }


                final int progreso = filaIndex;
                progressBar.setIndeterminate(false); // Asegura progreso real
                progressBar.setMaximum(totalFilas); // Por ejemplo, total de filas a procesar
                progressBar.setValue(progreso); // Esto lo vas actualizando por cada fila procesada
//                SwingUtilities.invokeLater(() -> progressBar.setValue(progreso));
            }

            logInfo("Proceso completado correctamente.");

        } catch (Exception e) {
            logError("Error durante el procesamiento: " + e.getMessage());
            e.printStackTrace();
        }
    }
    private int obtenerCantidadPreguntasValidas(Row filaEncuesta, Sheet sheetTemplate) {
        int indicePregunta = 0;

        while (true) {
            Cell celdaEncuesta = filaEncuesta.getCell(7 + indicePregunta); // desde columna H
            Row filaTemplate = sheetTemplate.getRow(indicePregunta + 1);   // desde fila 2


            Cell celdaTemplate = filaTemplate!=null?filaTemplate.getCell(1):null; // columna B
            if (celdaTemplate == null && celdaEncuesta == null) break;
            String valorEncuesta = obtenerValorCeldaComoTexto(celdaEncuesta).trim();
            String valorTemplate = obtenerValorCeldaComoTexto(celdaTemplate).trim();

            if (valorEncuesta.isEmpty() && valorTemplate.isEmpty()) break;

            if (!valorEncuesta.equals(valorTemplate)) {
                logError("Diferencia de pregunta en posición " + (indicePregunta + 1) + ": '" +
                        ((valorEncuesta==null||valorEncuesta=="")?"vacío":valorEncuesta) + "' ≠ '" +
                        ((valorTemplate==null||valorTemplate=="")?"vacío":valorTemplate) + "'");
                return -1;
            }

            indicePregunta++;
        }

        return indicePregunta;
    }


    private boolean filaEstaVacia(Row fila) {
        if (fila == null) return true;

        for (Cell celda : fila) {
            if (celda != null && !obtenerValorCeldaComoTexto(celda).trim().isEmpty()) {
                return false; // Hay al menos una celda con contenido útil
            }
        }
        return true;
    }

    private String obtenerValorCeldaComoTexto(Cell celda) {
        if (celda == null) return "";

        switch (celda.getCellType()) {
            case STRING:
                return celda.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(celda)) {
                    return celda.getDateCellValue().toString();
                } else {
                    return String.valueOf(celda.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(celda.getBooleanCellValue());
            case FORMULA:
                try {
                    return celda.getStringCellValue();
                } catch (IllegalStateException e) {
                    try {
                        return String.valueOf(celda.getNumericCellValue());
                    } catch (Exception ex) {
                        return "";
                    }
                }
            case BLANK:
            case _NONE:
            case ERROR:
            default:
                return "";
        }
    }

    private void copiarContenidoHoja(Sheet origen, Sheet destino) {
        HSSFWorkbook wbDestino = (HSSFWorkbook) destino.getWorkbook();
        HSSFWorkbook wbOrigen = (HSSFWorkbook) origen.getWorkbook();
        Map<Integer, CellStyle> estilosCopiados = new HashMap<>();

        for (int i = 0; i <= origen.getLastRowNum(); i++) {
            Row filaOrigen = origen.getRow(i);
            if (filaOrigen == null) continue;

            Row filaDestino = destino.createRow(i);
            filaDestino.setHeight(filaOrigen.getHeight());

            for (int j = 0; j < filaOrigen.getLastCellNum(); j++) {
                Cell celdaOrigen = filaOrigen.getCell(j);
                if (celdaOrigen == null) continue;

                Cell celdaDestino = filaDestino.createCell(j);
                copiarValorYEstiloCelda(wbOrigen, wbDestino, celdaOrigen, celdaDestino, estilosCopiados);
            }
        }

        // Copiar anchos de columna
        for (int i = 0; i <= origen.getRow(0).getLastCellNum(); i++) {
            destino.setColumnWidth(i, origen.getColumnWidth(i));
        }
    }


    private void copiarValorYEstiloCelda(HSSFWorkbook wbOrigen, HSSFWorkbook wbDestino,
                                         Cell celdaOrigen, Cell celdaDestino,
                                         Map<Integer, CellStyle> estilosCopiados) {
        // Copiar valor
        switch (celdaOrigen.getCellType()) {
            case STRING:
                celdaDestino.setCellValue(limpiarSaltosDeLinea(celdaOrigen.getStringCellValue()));
                break;
            case NUMERIC:
                celdaDestino.setCellValue(celdaOrigen.getNumericCellValue());
                break;
            case BOOLEAN:
                celdaDestino.setCellValue(celdaOrigen.getBooleanCellValue());
                break;
            case FORMULA:
                celdaDestino.setCellFormula(celdaOrigen.getCellFormula());
                break;
            case BLANK:
                celdaDestino.setBlank();
                break;
            default:
                break;
        }

        // Copiar estilo (caché por índice)
        CellStyle estiloOriginal = celdaOrigen.getCellStyle();
        if (estiloOriginal != null) {
            int hashCode = estiloOriginal.hashCode();
            CellStyle estiloDestino = estilosCopiados.get(hashCode);
            if (estiloDestino == null) {
                estiloDestino = wbDestino.createCellStyle();
                estiloDestino.cloneStyleFrom(estiloOriginal);
                estilosCopiados.put(hashCode, estiloDestino);
            }
            celdaDestino.setCellStyle(estiloDestino);
        }
    }


    private HSSFWorkbook copiarTemplate(HSSFWorkbook original) throws IOException {
        // Guardamos temporalmente el original en memoria
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        original.write(baos);
        ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
        return new HSSFWorkbook(bais);
    }
    // prueba gitflow
    // prueba hotfix
    private String convertirIndiceAColumnaExcel(int indice) {
        StringBuilder columna = new StringBuilder();
        while (indice >= 0) {
            columna.insert(0, (char) ('A' + (indice % 26)));
            indice = (indice / 26) - 1;
        }
        return columna.toString();
    }
    private static String limpiarSaltosDeLinea(String texto) {
        if (texto == null) return "";
        // Reemplaza saltos de línea múltiples por un espacio
        return texto.replaceAll("[\\r\\n]+", " ").trim();
    }

}
