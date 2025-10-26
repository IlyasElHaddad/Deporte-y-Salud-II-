public class Usuario {

    private String nombre;
    private int edad;
    private double peso;
    private double altura;

    public Usuario() {
        nombre = "Invitado";
        edad = 21;
        peso = 86;
        altura = 1.82;
    }

    public Usuario(String n, int e, double p, double a) {
        nombre = n;
        edad = e;
        peso = p;
        altura = a;
    }

    public String getNombre() { return nombre; }
    public void setNombre(String n) { nombre = n; }

    public int getEdad() { return edad; }
    public void setEdad(int e) { edad = e; }

    public double getPeso() { return peso; }
    public void setPeso(double p) { peso = p; }

    public double getAltura() { return altura; }
    public void setAltura(double a) { altura = a; }

    @Override
    public String toString() {
        return nombre + " (" + edad + " años, " + peso + " kg, " + altura + " m)";
    }
}
