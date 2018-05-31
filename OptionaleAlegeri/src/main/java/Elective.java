import java.util.ArrayList;

public class Elective {
    private String electiveName;
    private int maxStudents;
    private ArrayList<Student> students;

    public Elective(String name) {
        this.electiveName = name;
        students = new ArrayList<Student>();
    }

    public ArrayList<Student> getStudents() {
        return students;
    }
    public String getElectiveName() {
        return electiveName;
    }

    public void setElectiveName(String electiveName) {
        this.electiveName = electiveName;
    }

    public int getMaxStudents() {
        return maxStudents;
    }

    public void setMaxStudents(int maxStudents) {
        this.maxStudents = maxStudents;
    }

    public int addStudent(Student student) throws Exception {
        if (students.contains(student)) {
            throw new Exception("Duplicate student int elective: " + electiveName + " with name " + student.getName());
        }
        if (maxStudents > students.size()) {
            students.add(student);
            //student added succesfully to this elective
            return 0;
        } else {
            //cannot add more students in this elective
            return -1;
        }
    }
}
/*
* CellStyle blackStyle = workbook.createCellStyle();
            blackStyle.setFillForegroundColor(IndexedColors.BLACK.index);
            blackStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Sheet mySheet = workbook.createSheet("Repartizari");
            int i=0;
            int j=0;
            Row myRow = mySheet.createRow(i);
            myRow.createCell(j++).setCellValue("Nr.Crt.");
            myRow.createCell(j++).setCellValue("Nume");
            myRow.createCell(j++).setCellValue("Grupa");
            myRow.createCell(j++).setCellValue("Media");
            for(int k=1;k<=9;k++){
                myRow.createCell(j++).setCellValue(""+k);
            }
            myRow.createCell(j++).setCellStyle(blackStyle);
            for(int k=1;k<=5;k++){
                myRow.createCell(j++).setCellValue(""+k);
            }
            myRow.createCell(j++).setCellStyle(blackStyle);
            for(int k=1;k<=10;k++){
                myRow.createCell(j++).setCellValue(""+k);
            }*/
