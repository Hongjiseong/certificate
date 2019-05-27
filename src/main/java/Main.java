import service.Certification;
import service.Impl.TableCertification;

public class Main {
    public static void main(String[] args) {
        Certification certification = new TableCertification();
        certification.download("D:/tableExcel2.xlsx");
    }
}
