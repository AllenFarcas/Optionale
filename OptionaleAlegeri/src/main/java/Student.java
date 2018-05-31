import java.util.ArrayList;
import java.util.Comparator;

public class Student {
    private String name;
    private String group;
    private double media;
    private ArrayList<String> options;

    public Student () {
        options = new ArrayList<String>();
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGroup() {
        return group;
    }

    public void setGroup(String group) {
        this.group = group;
    }

    public double getMedia() {
        return media;
    }

    public void setMedia(double media) {
        this.media = media;
    }

    public ArrayList<String> getOptions() {
        return options;
    }

    public void addOption (String option) {
        options.add(option);
    }

    public static Comparator<Student> StuNameComparatorAscending = new Comparator<Student>() {

        public int compare(Student s1, Student s2) {
            String studentName1 = s1.getName().toUpperCase();
            String studentName2 = s2.getName().toUpperCase();
            //System.out.print("Student1: "+studentName1);
            //System.out.print(" Student2: "+studentName2);
            //System.out.println(" CompareToResult: "+studentName1.compareTo(studentName2));
            //ascending order
            //studentName1>studentName2 => >0;
            //studentName1<studentName2 => <0;
            return studentName1.compareTo(studentName2);

            //descending order
            //return StudentName2.compareTo(StudentName1);
        }
    };

    public static Comparator<Student> StuNameComparatorDescending = new Comparator<Student>() {

        public int compare(Student s1, Student s2) {
            String studentName1 = s1.getName().toUpperCase();
            String studentName2 = s2.getName().toUpperCase();
            //System.out.print("Student1: "+studentName1);
            //System.out.print(" Student2: "+studentName2);
            //System.out.println(" CompareToResult: "+studentName1.compareTo(studentName2));
            //ascending order
            //studentName1>studentName2 => >0;
            //studentName1<studentName2 => <0;
            //return studentName1.compareTo(studentName2);

            //descending order
            return studentName2.compareTo(studentName1);
        }
    };

    public static Comparator<Student> StuMediaComparatorDescending = new Comparator<Student>() {

        public int compare(Student s1, Student s2) {

            double media1 = s1.getMedia();
            double media2 = s2.getMedia();

            /*For ascending order
            if(media1-media2>=0) {
                //media1 >= media2
                return 1;
            } else {
                return -1;
            }*/
            /*For descending order*/
            //rollno2-rollno1;
            if(media1-media2<=0) {
                //media1 >= media2
                return 1;
            } else {
                return -1;
            }
        }
    };

    static Comparator<Student> StuMediaComparatorAscending = new Comparator<Student>() {

        public int compare(Student s1, Student s2) {

            double media1 = s1.getMedia();
            double media2 = s2.getMedia();

            //For ascending order
            if(media1-media2>=0) {
                //media1 >= media2
                return 1;
            } else {
                return -1;
            }
        }
    };
    @Override
    public String toString() {
        return name +" "+ group +" "+ media +" "+ options;
    }
}
