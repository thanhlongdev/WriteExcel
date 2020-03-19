import com.demo.hepler.WriteExcel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        Object[] objects1 = {"1", "Mickey"};
        Object[] objects2 = {"2", "Donal"};
        Object[] objects3 = {"3", "Jerry"};
        Object[] objects4 = {"4", "Tom"};
        List<Object[]> list = Arrays.asList(objects1, objects2, objects3, objects4);

        String path = "D:/book.xlsx";
        try{
            new WriteExcel().wirteExcel(list,path);
        } catch (IOException e){
            System.out.println(e.toString());
        }
    }
}
