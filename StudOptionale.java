import java.util.ArrayList;
import java.util.HashMap;
import java.util.Set;

public class StudOptionale {
    private double nrCrt;
    private String name;
    private String group;
    private double media;
    private HashMap<String,Integer> options;
    private String[] ex1Order;
    private String[] ex2Order;
    private String[] distrOrder;
    private HashMap<String, Boolean> alese;
    private HashMap<String, Boolean> sterse;
    private ArrayList<String[]> notAllowedEx1 = new ArrayList<>();

    public void initAleseSterse(){
        Set<String> auxSet = options.keySet();
        alese = new HashMap<>();
        sterse = new HashMap<>();
        for(String aux: auxSet){
            alese.put(aux,false);
            sterse.put(aux,false);
        }
    }
    public void stergeOpt(String optionala) throws Exception{
        if(alese.get(optionala)){
            alese.replace(optionala,false);
            sterse.replace(optionala,true);
        } else {
            throw new Exception ("Optionala: "+optionala+" exista in Alese cu FALSE la stud: "+name);
        }
    }
    public void resetAlese(){
        Set<String> auxSet = options.keySet();
        alese = new HashMap<>();
        for(String aux: auxSet){
            alese.put(aux,false);
        }
    }
    public void alegeOpt(String optionala) throws Exception{
        if(!alese.get(optionala)){
            if(!sterse.get(optionala)) {
                alese.replace(optionala, true);
            } else {
                throw new Exception ("Optionala: "+optionala+" a fost STEARSA si se incearca adaugare la stud: "+name);
            }
        } else {
            throw new Exception ("Optionala: "+optionala+" exista in Alese cu TRUE la stud: "+name);
        }
    }
    public boolean isAles(String optionala){
        return alese.get(optionala);
    }
    public boolean isSters(String optionala){
        return sterse.get(optionala);
    }
    public int getEx1Alese(){
        int count=0;
        //System.out.println("ex1Order: "+ex1Order);
        for(String exam:ex1Order){
            if(alese.get(exam)){
                count++;
            }
        }
        return count;
    }
    public int getEx2Alese(){
        int count=0;
        for(String exam:ex2Order){
            if(alese.get(exam)){
                count++;
            }
        }
        return count;
    }
    public int getDistrAlese(){
        int count=0;
        for(String exam:distrOrder){
            if(alese.get(exam)){
                count++;
            }
        }
        return count;
    }

    public void notAddEx1(String[] toAdd){
        notAllowedEx1.add(toAdd);
    }
    public boolean notEx1(){
        boolean val=true;
        for(String[] a:notAllowedEx1) {
            for(String x:a) {
                val &= isAles(x);
            }
        }
        return val;
    }

    public boolean getEx1Sterse(){
        int count=0;
        for(String exam:ex1Order){
            if(sterse.get(exam)){
                count++;
            }
        }
        return count==ex1Order.length;
    }
    public boolean getEx2Sterse(){
        int count=0;
        for(String exam:ex2Order){
            if(sterse.get(exam)){
                count++;
            }
        }
        return count==ex2Order.length;
    }
    public boolean getDistrSterse(){
        int count=0;
        for(String exam:distrOrder){
            if(sterse.get(exam)){
                count++;
            }
        }
        return count==distrOrder.length;
    }

    public StudOptionale () {
        options = new HashMap<String, Integer>();
    }

    public String[] getEx1Order() {
        return ex1Order;
    }

    public void setEx1Order(String[] ex1Order) {
        this.ex1Order = ex1Order;
    }

    public String[] getEx2Order() {
        return ex2Order;
    }

    public void setEx2Order(String[] ex2Order) {
        this.ex2Order = ex2Order;
    }

    public String[] getDistrOrder() {
        return distrOrder;
    }

    public void setDistrOrder(String[] distrOrder) {
        this.distrOrder = distrOrder;
    }

    public double getNrCrt() {
        return nrCrt;
    }

    public void setNrCrt(double nrCrt) {
        this.nrCrt = nrCrt;
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

    public HashMap<String, Integer> getOptions() {
        return options;
    }

    public void addOption (String option, int value) {
        options.put(option,value);
    }

    public String[] getOrderedElectives (HashMap<String,Integer> examene, ArrayList<String> exams) throws Exception{
        String[] exOrdine = new String[examene.size()];
        for(int k=1;k<=examene.size();k++){
            String subject = exams.get(k-1);
            int no = options.get(subject);
            if(no<0){
                throw new Exception("Problema la studentul: "+this.getName()+" materia: "+subject);
            } else {
                exOrdine[no - 1] = subject;
            }
        }
        //System.out.println("Student: "+name);
        //for(String o: exOrdine){
        //    System.out.print(" "+o);
        //}
        //System.out.println();
        return exOrdine;
    }

    @Override
    public String toString() {
        return "StudOptionale{" +
                "name='" + name + '\'' +
                ", group='" + group + '\'' +
                ", media=" + media +
                ", options=" + options +
                '}';
    }
}
