import java.util.Date;

public class Cell {

    private Date date;
    private double temperature;

    public Cell(Date date, double tempreature) {
        this.date = date;
        this.temperature = tempreature;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public double getTempreature() {
        return temperature;
    }

    public void setTempreature(double tempreature) {
        this.temperature = tempreature;
    }

    @Override
    public String toString() {
        return "Cell{" +
                "date=" + date +
                ", tempreature=" + temperature +
                '}';
    }
}
