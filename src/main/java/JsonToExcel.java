import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Paths;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class JsonToExcel {
    //map to hold temp data to be written to
    HashMap<String, Object> tempRowData = new HashMap<>();

    //array list to store hash list of maps of rows
    ArrayList<HashMap<String,Object>> rowList = new ArrayList<>();

    public static HashMap<String, Object> copyMap(HashMap<String, Object> original) {
        HashMap<String, Object> copy = new HashMap<String,Object>();
        for (Map.Entry<String, Object> entry : original.entrySet()){
            copy.put(entry.getKey(), entry.getValue());
        }
        return copy;
    }
    public String makePath(String pathKey,String key){

        if(pathKey == null){
            if(key.contains(" ")){
                return  "["+key+"]";
            }else{
                return key;
            }
        }else if(key.contains(" ")){
            return pathKey + "["+"\""+key+"\""+"]";
        }else{
            return pathKey + "."+key;
        }

    }


    //recursive method for reading template and updating data with excel value
    public void hashMapper(Map<String, Object> lhm1,String parentKey){
        for (Map.Entry<String, Object> entry : lhm1.entrySet()) {

            String key = entry.getKey();
            Object value = entry.getValue();
            if (handelInstance(value) >= 0) {
                value = handelGeneration(key,value);
                tempRowData.put(makePath(parentKey,key),value);
            } else if (value instanceof Map) {
                Map<String, Object> subMap = (Map<String, Object>)value;
                hashMapper(subMap,makePath(parentKey,key));
            }else if (value instanceof Iterator<?> || value instanceof ArrayList<?>){
                Iterator<?> name = ((Iterable<?>) value).iterator();
                //loop thorough nested json
                int iteratorIndex = 0;
                while(name.hasNext()){

                    Object val = name.next();
                   try {
                       hashMapper((Map<String, Object>)val,makePath(parentKey,key));
                   }catch (ClassCastException e){
                       //if you can't cast value to map, store directly
                       tempRowData.put(makePath(parentKey,key+"["+iteratorIndex+"]"),val);
                   }
                    iteratorIndex+=1;
               }
            }else if(value ==  null){
                value = handelGeneration(key,value);
                tempRowData.put(makePath(parentKey,key),value);
            }else{
                tempRowData.put(makePath(parentKey,key),value);
                System.out.println("could not determine instance for :"+makePath(parentKey,key)+"-*-"+key+"\n"+value);
            }
        }
    }

    //handel dynamic value generation
    private Object handelGeneration(String key, Object value) {

        return value;
    }

    //a method for writing json string into jaa object
    public void exportJsonToJavaObject(String filePath){

        ArrayList<String> listOfJsonFiles = listFilesFromDir(filePath);

        for (int i = 0; i < listOfJsonFiles.size(); i++) {

            String fileName = filePath + "/" + listOfJsonFiles.get(i);
            mapToRowMap(jsonToFlatMap(fileName,false));

        }
    }

    //function to convert json to map
    public List<Map<String, Object>> jsonToFlatMap(String fileName,Boolean template){
        ObjectMapper mapper = new ObjectMapper();

        //map for array encounter
        List<Map<String, Object>> map = new ArrayList<>();
        //map for json object encounter
        Map<String, Object> map2 = new HashMap<>();
        // convert JSON file to map
        try {
            map = mapper.readValue(Paths.get(fileName).toFile(),  new TypeReference<List<Map<String, Object>>>(){});
        } catch (IOException e) {
            try {
                map2 = mapper.readValue(Paths.get(fileName).toFile(),  new TypeReference<Map<String, Object>>(){});
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
        if(!map.isEmpty()){
            return map;
        }else{


            //check depth for array
            //int depth = 0;
            int keySize = map2.size();
            int keySetSize = map2.keySet().size();
            for (String key: map2.keySet()) {

                if(keySize == 1 &&  keySetSize == 1 &&  (map2.get(key) instanceof Iterable<?> || map2.get(key) instanceof ArrayList<?>)){
                    Iterator<?> name = ((Iterable<?>) map2.get(key)).iterator();
                    while (name.hasNext()){
                        try{
                            map.add((Map<String, Object>)name.next());
                        }catch(ClassCastException e){
                            map.add(map2);
                            break;
                        }
                    }
                    break;
                }else{

                     map.add(map2);
                     break;
                }
            }
            return map;
        }
    }


    //function to loop through each item in json array as a raw
    public void mapToRowMap(List<Map<String,Object>> map){

        for (int i = 0; i < map.size(); i++) {
            //recursive check if each item in a json arr
            hashMapper(map.get(i),"x["+i+"]");
            //update the row list with new map
            rowList.add(copyMap(tempRowData));
            tempRowData.clear();
        }
        //Write result to excel
        writeMapToExcel(rowList,"./src/main/resources/results/");
        rowList.clear();
    }



    //function to write rows to excel
    private void writeMapToExcel(ArrayList<HashMap<String, Object>> rowList,String pathString) {

        String fileName =  new SimpleDateFormat("yyyyMMddHHmmssSS").format(Calendar.getInstance().getTime());
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("data_"+fileName);

        int rowCount = -1;
        Set<String> keySets =  rowList.get(0).keySet();
        for (int i = 0; i < rowList.size(); i++) {

            int columnCount = -1;
            Row row = sheet.createRow(++rowCount);

            if(rowCount == 0){
                for (String header: keySets) {
                    Cell cell = row.createCell(++columnCount);
                    cell.setCellValue(header);
                }
                columnCount -= 1;
                i-=1;
                continue;
            }
            for (String key: keySets) {
                Cell cell = row.createCell(++columnCount);

                if(rowList.get(i).containsKey(key)){
                    Object value = rowList.get(i).get(key);
                    int objType = handelInstance(value);

                    switch (objType){
                        case 0:
                            cell.setCellValue((Integer)value);
                            break;
                        case 1:
                            cell.setCellValue((Double) value);
                            break;
                        case 2:
                            cell.setCellValue((Float) value);
                            break;
                        case 3:
                            cell.setCellValue((Long) value);
                            break;
                        case 4:
                            cell.setCellValue((Date) value);
                            break;
                        case 5:
                            cell.setCellValue((Boolean) value);
                            break;
                        case 6:
                            cell.setCellValue((Time) value);
                            break;
                        case 7:
                            cell.setCellValue((Timestamp)value);
                            break;
                        case 8:
                            cell.setCellValue((String)value);
                        case -1:
                            cell.setCellValue("");
                        default:
                            cell.setCellValue((String)value);
                            break;
                    }
                }
            }
        }
        //write to excel
        FileOutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(pathString + fileName + ".xlsx");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    //function to return list of files
    public ArrayList<String> listFilesFromDir(String dir){
        File folder = new File(dir);
        File[] listOfFiles = folder.listFiles();
        ArrayList<String> listOfGoodFiles = new ArrayList<>();

        for (int i = 0; i < listOfFiles.length; i++) {
            if (listOfFiles[i].isFile()) {
                listOfGoodFiles.add(listOfFiles[i].getName());
            } else if (listOfFiles[i].isDirectory()) {
                System.out.println("Directory Encountered While searching for Json: " + listOfFiles[i].getName());
            }
        }
        return listOfGoodFiles;
    }
    //handel object instances
    public static int handelInstance(Object obj){
        if(obj instanceof Integer) return 0;
        if(obj instanceof Double)  return 1;
        if(obj instanceof Float)  return 2;
        if(obj instanceof Long)  return 3;
        if(obj instanceof Date)  return 4;
        if(obj instanceof Boolean)  return 5;
        if(obj instanceof Time)  return 6;
        if(obj instanceof Timestamp)  return 7;
        if(obj instanceof String)  return 8;
        if(obj == null)  return -1;
        return -2;
    }

    //main method
    public static void main(String[] args) throws IOException, ParseException {
        //file path to root folder of the input/output folder
        String dataFilePath = "./src/main/resources/jsonTemplate";
        JsonToExcel jTE = new JsonToExcel();
        jTE.exportJsonToJavaObject(dataFilePath);


    }


}
