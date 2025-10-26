public class Entrenamiento {

    private String tipo;
    private String intensidad;
    private int duracion;
    private int calorias;

    public Entrenamiento() {
        tipo = "Sin definir";
        intensidad = "Moderado";
        duracion = 0;
        calorias = 0;
    }

    public Entrenamiento(String t, String i, int d, int c) {
        tipo = t;
        intensidad = i;
        duracion = d;
        calorias = c;
    }

    public String getTipo() { return tipo; }
    public void setTipo(String t) { tipo = t; }

    public String getIntensidad() { return intensidad; }
    public void setIntensidad(String i) { intensidad = i; }

    public int getDuracion() { return duracion; }
    public void setDuracion(int d) { duracion = d; }

    public int getCalorias() { return calorias; }
    public void setCalorias(int c) { calorias = c; }

    @Override
    public String toString() {
        return tipo + " - " + intensidad + " (" + duracion + " min, " + calorias + " kcal)";
    }
}
