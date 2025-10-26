import java.awt.*;
import javax.swing.*;

public class ElegirEntreno extends JFrame {

    private Entrenamiento seleccionado;
    private JLabel lblResumen;
    private JComboBox<String> comboHorario;
    private JButton btnConfirmar;
    private JButton btnVolver;

    public ElegirEntreno(Entrenamiento e) {
        this.seleccionado = e;
        initComponentsGridBag();
        setSize(420, 260);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
    }

    private void initComponentsGridBag() {
        setLayout(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();

        lblResumen = new JLabel(seleccionado.toString());
        lblResumen.setPreferredSize(new Dimension(360, 60));

        comboHorario = new JComboBox<>();
        comboHorario.addItem("Mañana 07:00");
        comboHorario.addItem("Mediodía 12:00");
        comboHorario.addItem("Tarde 18:00");
        comboHorario.addItem("Noche 21:00");

        btnConfirmar = new JButton("Confirmar");
        btnConfirmar.addActionListener(e -> confirmar());

        btnVolver = new JButton("Volver");
        btnVolver.addActionListener(e -> dispose());

        c.insets = new Insets(8,8,8,8);
        c.gridx = 0; c.gridy = 0; c.gridwidth = 2; c.fill = GridBagConstraints.HORIZONTAL;
        add(lblResumen, c);

        c.gridwidth = 1; c.gridy = 1; c.gridx = 0; c.weightx = 0.3;
        add(new JLabel("Horario:"), c);
        c.gridx = 1; c.weightx = 0.7; c.fill = GridBagConstraints.HORIZONTAL;
        add(comboHorario, c);

        c.gridy = 2; c.gridx = 0; c.weightx = 0.5; c.fill = GridBagConstraints.NONE;
        add(btnConfirmar, c);
        c.gridx = 1;
        add(btnVolver, c);
    }

    private void confirmar() {
        String horario = (String) comboHorario.getSelectedItem();
        String mensaje = String.format("Entrenamiento confirmado:\n%s\nHorario: %s",
                seleccionado.toString(), horario);
        JOptionPane.showMessageDialog(this, mensaje);
        dispose();
    }
}
