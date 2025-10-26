import java.awt.*;
import java.awt.event.*;
import java.net.URL;
import java.util.ArrayList;
import javax.swing.*;

public class AppDeporte extends JFrame {

    private JPanel panelIzquierda;
    private JPanel panelCentro;
    private JPanel panelDerecha;
    private JComboBox<String> comboFiltro;
    private DefaultListModel<Entrenamiento> modeloLista;
    private JList<Entrenamiento> listaEntrenos;
    private JButton btnElegir;
    private JButton btnSalir;
    private JLabel lblImagenLocal;
    private JLabel lblImagenInternet;
    private JButton btnExportar;

    private java.util.List<Entrenamiento> lista = new ArrayList<>();

    public AppDeporte() {
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(900, 520);
        setLocationRelativeTo(null);
        cargarDatos();
        initgridbag();
    }

    private void cargarDatos() {
        lista.add(new Entrenamiento("HIIT", "Intenso", 30, 450));
        lista.add(new Entrenamiento("Cardio suave", "Relajado", 45, 300));
        lista.add(new Entrenamiento("Fuerza pierna", "Moderado", 50, 400));
        lista.add(new Entrenamiento("Core", "Mixto", 25, 180));
        lista.add(new Entrenamiento("Spinning", "Intenso", 40, 500));
        lista.add(new Entrenamiento("Yoga Vinyasa", "Relajado", 60, 220));
    }

    private void initgridbag() {
        getContentPane().setLayout(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();

        panelIzquierda = new JPanel(new BorderLayout());
        lblImagenLocal = new JLabel("", SwingConstants.CENTER);
        ImagenLocal("/img/Deporte.png", lblImagenLocal);
        panelIzquierda.add(lblImagenLocal, BorderLayout.CENTER);

        panelDerecha = new JPanel(new BorderLayout());
        lblImagenInternet = new JLabel("",SwingConstants.CENTER);
        ImagenInternet("https://previews.123rf.com/images/oksancia/oksancia1212/oksancia121200168/16759357-sports-balls-vertical-seamless-pattern-background-border.jpg", lblImagenInternet);
        panelDerecha.add(lblImagenInternet, BorderLayout.CENTER);


        panelCentro = new JPanel(new GridBagLayout());
        GridBagConstraints b = new GridBagConstraints();

        comboFiltro = new JComboBox<>();
        comboFiltro.addItem("Todos");
        comboFiltro.addItem("Intenso");
        comboFiltro.addItem("Moderado");
        comboFiltro.addItem("Relajado");
        comboFiltro.addItem("Mixto");

        modeloLista = new DefaultListModel<>();
        listaEntrenos = new JList<>(modeloLista);
        JScrollPane scrollLista = new JScrollPane(listaEntrenos);
        scrollLista.setPreferredSize(new Dimension(250, 200));

        btnElegir = new JButton("Elegir entrenamiento");
        btnExportar = new JButton("Exportar a Excel");
        btnSalir = new JButton("Salir");

        comboFiltro.addActionListener(e -> actualizarLista());
        btnElegir.addActionListener(e -> abrirElegirEntreno());
        btnSalir.addActionListener(e -> dispose());
        btnExportar.addActionListener(e -> exportarEntreno());

        b.insets = new Insets(10, 10, 10, 10);
        b.gridx = 0; b.gridy = 0;
        b.anchor = GridBagConstraints.WEST;
        panelCentro.add(new JLabel("Filtrar por intensidad:"), b);
        b.gridx = 1;
        b.fill = GridBagConstraints.HORIZONTAL;
        b.weightx = 1.0;
        panelCentro.add(comboFiltro, b);

        b.gridx = 0; b.gridy = 1; b.gridwidth = 2;
        b.fill = GridBagConstraints.BOTH;
        b.weightx = 1.0;  b.weighty = 1.0;
        panelCentro.add(scrollLista, b);

        b.gridy = 2; b.gridwidth = 1; b.weighty = 0;
        b.fill = GridBagConstraints.BOTH;
        b.weightx = 0.5;
        panelCentro.add(btnElegir, b);

        b.gridx = 1; b.weightx = 0.5;
        panelCentro.add(btnSalir, b);
        
        b.gridy = 3; b.gridx = 0; b.gridwidth = 2;
        b.fill = GridBagConstraints.HORIZONTAL; b.weightx = 1.0; b.weighty = 0;
        panelCentro.add(btnExportar, b);

        c.insets = new Insets(10,10, 10, 10);

        c.gridx = 0; c.gridy = 0; c.weightx = 0.3; c.weighty = 1.0;
        c.fill = GridBagConstraints.BOTH;
        getContentPane().add(panelIzquierda, c);

        c.gridx = 1; c.gridy = 0; c.weightx = 0.4;
        getContentPane().add(panelCentro, c);

        c.gridx = 2; c.gridy = 0; c.weightx = 0.3;
        getContentPane().add(panelDerecha, c);
        actualizarLista();
        addComponentListener(new ComponentAdapter() {
            @Override
            public void componentResized(ComponentEvent e) {
                redimensionarImagen(lblImagenLocal);
                redimensionarImagen(lblImagenInternet);
                panelIzquierda.setPreferredSize(new Dimension(200, 200));
                lblImagenLocal.setPreferredSize(new Dimension(200, 200));
                panelDerecha.setPreferredSize(new Dimension(200, 200));
                lblImagenInternet.setPreferredSize(new Dimension(200, 200));}
        });
    }

    private void actualizarLista() {
        modeloLista.clear();
        String filtro = (String) comboFiltro.getSelectedItem();
        for (Entrenamiento e : lista) {
            if (filtro.equals("Todos") || e.getIntensidad().equalsIgnoreCase(filtro)) {
                modeloLista.addElement(e);
            }
        }
    }

    private void abrirElegirEntreno() {
        Entrenamiento e = listaEntrenos.getSelectedValue();
        if (e == null) {
            JOptionPane.showMessageDialog(this, "Selecciona un entrenamiento primero.");
            return;
        }
        ElegirEntreno ventana = new ElegirEntreno(e);
        ventana.setVisible(true);
    }

    private void ImagenLocal(String ruta, JLabel lbl) {
    try {java.net.URL imgUrl = getClass().getResource(ruta);
        if (imgUrl != null) {
            ImageIcon icon = new ImageIcon(imgUrl);
            lbl.setIcon(icon);
        } else {}
    } catch (Exception ex) {}
    }
    private void ImagenInternet(String url, JLabel lbl) {
        new Thread(() -> {
            try {
                ImageIcon icon = new ImageIcon(new URL(url));
                SwingUtilities.invokeLater(() -> lbl.setIcon(icon));
            } catch (Exception ex) {}
        }).start();
    }

    private void redimensionarImagen(JLabel lbl) {
        Icon icono = lbl.getIcon();
        if (icono == null || !(icono instanceof ImageIcon)) return;

        Image img = ((ImageIcon) icono).getImage();
        int w = lbl.getWidth();
        int h = lbl.getHeight();
        if (w > 0 && h > 0) {
            Image escalada = img.getScaledInstance(w, h, Image.SCALE_SMOOTH);
            lbl.setIcon(new ImageIcon(escalada));
        }
    }
    private void exportarEntreno() {
    Entrenamiento seleccionado = listaEntrenos.getSelectedValue();
    if (seleccionado == null) {
        JOptionPane.showMessageDialog(this, "Selecciona un entrenamiento.");
        return;
    }

    java.util.List<Entrenamiento> unoSolo = new ArrayList<>();
    unoSolo.add(seleccionado);

    ExcelEntrenamientos.escribirEntrenamientos(unoSolo);

    JOptionPane.showMessageDialog(this, 
        "Entrenamiento guardado 'Entrenamientos.xlsx'");
}
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            new AppDeporte().setVisible(true);
        });
    }
}
