import com.fasterxml.jackson.databind.ObjectMapper;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Stream;

public class ExcelToJson {
    /************************************************************************************
    *                   For processing Excel records to JsonObject                       *
    *                  Uses Json template from Resources/JsonTemplates folder            *
    *                   And Updates the json records based on the Excel value            *
    *                                                                                    *
    ***************************************************************************************/
    private  ArrayList<String> pathsToGenerateVal = new ArrayList<>();

    JsonToExcel jtee = new JsonToExcel();
    //to hold list of iterable
    ArrayList<Object> listOfIterable = new ArrayList<>();
    private boolean iterable = false;
    private String iterableKey;

    public void excelToJsonMapper(String templatePath, String excelPath, String outputPath) throws IOException {

        //object
        JsonToExcel jt = new JsonToExcel();

        ArrayList<String> listOfExcelFiles = jt.listFilesFromDir(excelPath);
        ArrayList<String> listOfJsonTemplate = jt.listFilesFromDir(templatePath);

        for (int i = 0; i < listOfExcelFiles.size(); i++) {

            ArrayList<HashMap<String, Object>> tempExcelDataRow = new ArrayList<>();
            tempExcelDataRow = ReadExcel.readExcelFromPath(excelPath + "/" + listOfExcelFiles.get(i));
            for (int j = 0; j < listOfJsonTemplate.size(); j++) {

                HashMap<String, Object> tempJsonRaw =new HashMap<>();
                List<Map<String, Object >> tempJsonTemplateRow;
                tempJsonTemplateRow = jt.jsonToFlatMap(templatePath+"/"+listOfJsonTemplate.get(j),true);
                //hashMapper(tempJsonTemplateRow.get(0),tempJsonRaw,false);

                JsonToExcel jte = new JsonToExcel();
                jte.hashMapper(tempJsonTemplateRow.get(0),"x["+0+"]");
                tempJsonRaw = jte.tempRowData;

                if(tempExcelDataRow.get(0).keySet().containsAll(tempJsonRaw.keySet()) &&
                        tempJsonRaw.keySet().containsAll(tempExcelDataRow.get(0).keySet())){

                    for (int l = 0; l < tempExcelDataRow.size(); l++) {
                        hashMapper(tempJsonTemplateRow.get(0),tempExcelDataRow.get(l),"x["+0+"]");
                        //write each rom from excel to json
                        writeDataToJson(outputPath + "/" + listOfJsonTemplate.get(j).replace(".json","")+l+"_",tempJsonTemplateRow);
                    }
                    tempJsonRaw.clear();
                    tempJsonTemplateRow.clear();
                    tempExcelDataRow.clear();
                    break;
                }else{
                    tempJsonRaw.clear();
                    tempJsonTemplateRow.clear();
                }

            }
            tempExcelDataRow.clear();
        }
    }
    //recursive method for reading template and updating data with excel value
    public void hashMapper(Map<String, Object> lhm1,HashMap<String, Object> tempRaw,String parentKey){
        for (Map.Entry<String, Object> entry : lhm1.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();
            if (JsonToExcel.handelInstance(value) >= 0) {

                lhm1.put(key,updateHelper(jtee.makePath(parentKey,key),tempRaw));

            } else if (value instanceof Map) {

                Map<String, Object> subMap = (Map<String, Object>)value;
                hashMapper(subMap,tempRaw,jtee.makePath(parentKey,key));
            }else if (value instanceof Iterator<?> || value instanceof ArrayList<?>) {

                if(!this.listOfIterable.isEmpty()){
                    if(!this.iterable){
                        this.listOfIterable.clear();
                    }
                }
                if(value instanceof ArrayList<?> && !this.iterable  && this.listOfIterable.isEmpty()){
                    this.iterableKey = key;
                    listOfIterable.add(value);
                    this.iterable = true;
                }
                Iterator<?> name = ((Iterable<?>) value).iterator();
                int iteratorIndex = 0;
                while(name.hasNext()){
                    Object val = name.next();
                    try {
                        hashMapper((Map<String, Object>)val,tempRaw,jtee.makePath(parentKey,key));
                    }catch (ClassCastException e){


                            if(!this.iterable) continue;

                            if(this.iterable && !this.listOfIterable.isEmpty()){
                                lhm1.put(key,this.listOfIterable.get(0));
                                this.iterable = false;
                            }else{

                                lhm1.put(key,updateHelper(jtee.makePath(parentKey,key+iteratorIndex),tempRaw));
                            }
                    }
                    iteratorIndex+=1;
                }
            }else if(value == null){
                System.out.println("Null Value Detected for: "+key);
                lhm1.put(key,updateHelper(jtee.makePath(parentKey,key),tempRaw));

            }else{

                lhm1.put(key,updateHelper(jtee.makePath(parentKey,key),tempRaw));
                System.out.println(key+" could not determine instance for :"+value);
            }
        }
    }

    //write excel to json
    public void writeDataToJson(String filePath, List<Map<String, Object >> data) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        //writing the resut to json file
        String fileName =  new SimpleDateFormat("yyyyMMddHHmmss_SS").format(Calendar.getInstance().getTime());
        mapper.writerWithDefaultPrettyPrinter()
                .writeValue(Paths.get(filePath+fileName+".json").toFile(), data.get(0));
    }
    //handel dynamic path-string
    public boolean dynamicPathString(String src, String dst){
        return dst.contains(src);
    }

    //check if key in updatable path
    public boolean compareToGenerate(String key){
        for (int i = 0; i < pathsToGenerateVal.size(); i++) {
            if(dynamicPathString(key,pathsToGenerateVal.get(i))){
                return true;
            }
        }
        return false;
    }

    //add key to the updatable path
    public void addToUpdatable(String file){
        try (BufferedReader br = new BufferedReader(new FileReader(file))) {
            String line;
            while ((line = br.readLine()) != null) {
                this.pathsToGenerateVal.add(line);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    //helper function to update json calue
    public Object updateHelper(String key,HashMap<String, Object> tempRaw){

        //System.out.println("tempRaw key set: "+tempRaw.keySet().toString());
        System.out.println("key:"+key);
        if(tempRaw.containsKey(key)){
            if((key.contains("uuid") || key.contains("id"))
                    && compareToGenerate(key)){
                return UUID.randomUUID().toString();
            }else if(key.contains("date") && compareToGenerate(key)){
                return new SimpleDateFormat("yyyy.MM.dd").format(new Date());
            }else if(key.contains("timeStamp") && compareToGenerate(key)){
                return new SimpleDateFormat("yyyy.MM.dd.HH.mm.ss.SS").format(new Date());
            }
        }

        return tempRaw.get(key);
    }

    //main method
    public static void main(String[] args) throws IOException, ParseException {
        //file path to root folder of the input/output folder
        String templateFilePath = "./src/main/resources/jsonTemplate";
        String excelFilePath = "./src/main/resources/results";
        String jsonResultPath = "./src/main/resources/jsonResult";
        String pathToGenerationList = "./src/main/resources/genResult/path.txt";

        ExcelToJson eTJ = new ExcelToJson();
        eTJ.addToUpdatable(pathToGenerationList);
        eTJ.excelToJsonMapper(templateFilePath,excelFilePath,jsonResultPath);

    }



}
