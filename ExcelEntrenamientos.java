import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.IOUtils;
import java.io.*;
import java.util.List;

public class ExcelEntrenamientos {

    private static final String NOMBRE_ARCHIVO = "Entrenamientos.xlsx";


    public static void escribirEntrenamientos(List<Entrenamiento> entrenos) {
        Workbook libro;
        Sheet hoja;

        File archivo = new File(NOMBRE_ARCHIVO);

        try {
            if (archivo.exists()) {
                try (FileInputStream fis = new FileInputStream(archivo)) {
                    libro = new XSSFWorkbook(fis);
                }
                hoja = libro.getSheet("Entrenamientos");
                if (hoja == null) hoja = libro.createSheet("Entrenamientos");
            } else {
                libro = new XSSFWorkbook();
                hoja = libro.createSheet("Entrenamientos");

                Row cabecera = hoja.createRow(0);
                String[] titulos = {"Tipo", "Intensidad", "Duración (min)", "Calorías"};
                CellStyle estiloCabecera = libro.createCellStyle();
                Font fuente = libro.createFont();
                fuente.setBold(true);
                fuente.setFontHeightInPoints((short) 12);
                estiloCabecera.setFont(fuente);
                estiloCabecera.setAlignment(HorizontalAlignment.CENTER);
                for (int i = 0; i < titulos.length; i++) {
                    Cell celda = cabecera.createCell(i);
                    celda.setCellValue(titulos[i]);
                    celda.setCellStyle(estiloCabecera);
                }
            }

            for (Entrenamiento e : entrenos) {
                int fila = buscarFilaPorTipo(hoja, e.getTipo());
                if (fila == -1) {
                    fila = hoja.getLastRowNum() + 1;
                }
                Row r = hoja.getRow(fila);
                if (r == null) r = hoja.createRow(fila);

                r.createCell(0).setCellValue(e.getTipo());
                r.createCell(1).setCellValue(e.getIntensidad());
                r.createCell(2).setCellValue(e.getDuracion());
                r.createCell(3).setCellValue(e.getCalorias());
            }

            for (int i = 0; i < 4; i++) hoja.autoSizeColumn(i);

            if (!tieneImagen(hoja)) {
                insertarImagen(libro, hoja, "/img/Deporte.png", hoja.getLastRowNum() + 2, 0);
            }

            try (FileOutputStream fos = new FileOutputStream(archivo)) {
                libro.write(fos);
            }

            System.out.println("Excel actualizado.");

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private static int buscarFilaPorTipo(Sheet hoja, String tipo) {
        for (int i = 1; i <= hoja.getLastRowNum(); i++) {
            Row r = hoja.getRow(i);
            if (r != null && r.getCell(0) != null) {
                if (r.getCell(0).getStringCellValue().equalsIgnoreCase(tipo)) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static boolean tieneImagen(Sheet hoja) {
        if (hoja.getDrawingPatriarch() == null) return false;
        return hoja.getDrawingPatriarch().iterator().hasNext();
    }
    private static void insertarImagen(Workbook libro, Sheet hoja, String ruta, int fila, int col) {
        try (InputStream is = new FileInputStream(ruta)) {
            byte[] bytes = IOUtils.toByteArray(is);
            int idx = libro.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
            CreationHelper helper = libro.getCreationHelper();
            Drawing<?> drawing = hoja.createDrawingPatriarch();
            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setCol1(col);
            anchor.setRow1(fila);
            Picture pict = drawing.createPicture(anchor, idx);
            pict.resize(3.0);
        } catch (IOException ex) {}
    }
}
