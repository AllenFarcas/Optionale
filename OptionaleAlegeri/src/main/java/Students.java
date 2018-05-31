import java.util.ArrayList;
import java.util.Collections;

public class Students {
    ArrayList<Student> studentList;
    public Students (){
        studentList = new ArrayList<Student>();
    }
    public ArrayList<Student> getStudentList() {
        return studentList;
    }
    public void addStudent (Student student){
        studentList.add(student);
    }
    public void sortNameAscending () {
        System.out.println("Student Name Sorting:");
        Collections.sort(studentList, Student.StuNameComparatorAscending);
        //printStudentList();
    }
    public void sortNameDescending () {
        System.out.println("Student Name Sorting:");
        Collections.sort(studentList, Student.StuNameComparatorDescending);
        //printStudentList();
    }
    public void sortMediaAscending () {
        System.out.println("Student Media Sorting:");
        Collections.sort(studentList, Student.StuMediaComparatorAscending);
        //printStudentList();
    }
    public void sortMediaDescending () {
        System.out.println("Student Media Sorting:");
        Collections.sort(studentList, Student.StuMediaComparatorDescending);
        //printStudentList();
    }
    public void printStudentList(){
        for (Student str : studentList) {
            System.out.println(str.toString());
        }
    }

    @Override
    public String toString() {
        return studentList.toString();
    }
}
