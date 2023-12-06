package com.enssel.excel.file;

import com.grapecity.documents.excel.O;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * .xlsx To xml
 * <p>
 * 엑셀데이터를 zip으로 변환 후 xml 데이터를 기반으로 엑셀데이터를 추출하는 작업을 한다.
 */
public class XTX {
    private final File excelFile;
    private final File zipFile;
    private final File unzipFolder;

    private Object[] data;

    public XTX(String filePath) {
        String tempFileName = UUID.randomUUID().toString();

        System.out.println("생성된 파일명" + tempFileName);

        excelFile = new File(filePath);
        zipFile = new File("./tempFile/" + tempFileName + ".zip");
        unzipFolder = new File(zipFile.getParent(), zipFile.getName().replace(".zip", ""));

        // .zip 생성
        createZipFile();

        // 압축 해제
        unzipFile();

        readXml();

        // 임시 데이터 삭제
        if (unzipFolder.exists()) deleteFile(unzipFolder);
        if (zipFile.exists()) deleteFile(zipFile);
    }

    private void createZipFile() {
        // .zip 생성
        try {
            // 파일 경로에 파일이 없을 경우 생성
            File folder = new File("./tempFile");
            if (!folder.exists()) {
                folder.mkdir();
            }

            FileInputStream input = new FileInputStream(excelFile);
            FileOutputStream output = new FileOutputStream(zipFile);

            byte[] buf = new byte[1024];

            int readData;

            while ((readData = input.read(buf)) > 0) {
                output.write(buf, 0, readData);
            }

            input.close();
            output.close();
        } catch (Exception e) {
            System.out.println("오류발생");
            e.printStackTrace();
        }
    }

    private void unzipFile() {


        if (!unzipFolder.exists()) unzipFolder.mkdir();

        try (
                FileInputStream fis = new FileInputStream(zipFile);
                ZipInputStream zis = new ZipInputStream(fis);
                BufferedInputStream bis = new BufferedInputStream(zis);
        ) {
            ZipEntry zipEntry = null;
            while ((zipEntry = zis.getNextEntry()) != null) {
                File f = new File(unzipFolder.getAbsolutePath(), zipEntry.getName());
                if (zipEntry.isDirectory()) {
                    f.mkdir();
                } else {
                    if (!f.getParentFile().exists()) f.getParentFile().mkdir();
                    try (
                            FileOutputStream fos = new FileOutputStream(f);
                            BufferedOutputStream bos = new BufferedOutputStream(fos);
                    ) {
                        int b = 0;
                        while ((b = bis.read()) != -1) {
                            bos.flush();
                            bos.write(b);
                        }
                    }
                }
            }

        } catch (Exception e) {
            System.out.println("압축해제 중 오류 발생");
            e.printStackTrace();
        }
    }

    /**
     * 파일 삭제
     *
     * @param file 삭제할 파일이나 디렉토리
     */
    private static void deleteFile(File file) {
        try {
            if (file.exists()) {
                // 파일이 디렉토리의 경우 내부 파일을 전부 제거 해야 한다.
                if (file.isDirectory()) {
                    File[] folder_list = file.listFiles(); //파일리스트 얻어오기

                    for (int i = 0; i < folder_list.length; i++) {
                        if (folder_list[i].isFile()) folder_list[i].delete();
                        else deleteFile(folder_list[i]); //재귀함수호출

                        folder_list[i].delete();
                    }
                }

                file.delete(); //파일(폴더) 삭제
            }
        } catch (Exception e) {
            e.getStackTrace();
        }
    }

    private void readXml() {
        String path = unzipFolder.getPath() + "/xl/pivotCache";
        File[] pivotCaches = new File(path).listFiles();

        if (pivotCaches == null) {
            System.out.println("피벗 데이터가 없는 파일 입니다.");
            return;
        }

        try {
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();


            int cacheCount = 0;

            for (File f : pivotCaches) {
                if (f.isFile() && f.getName().startsWith("pivotCacheDefinition")) cacheCount++;
            }

            Object[] reaultData = new Object[cacheCount];

            for (int i = 0; i < cacheCount; i++) {
                Document pivotCacheDefinition = builder.parse(path + "/pivotCacheDefinition" + (i + 1) + ".xml");
                Element pivotCacheDefinitionElement = pivotCacheDefinition.getDocumentElement();

                NodeList definitionElement = pivotCacheDefinitionElement.getChildNodes();

                Map<String, List<String>> definitionMap = new LinkedHashMap<>();

                for (int cn = 0, n1 = definitionElement.getLength(); cn < n1; cn++) {
                    Node item = definitionElement.item(cn);

                    if (item.getNodeType() == Node.ELEMENT_NODE && "cacheFields".equals(item.getNodeName())) {
                        NodeList cacheFieldList = item.getChildNodes();

                        for (int cfi = 0, cfl = cacheFieldList.getLength(); cfi < cfl; cfi++) {
                            Node cf = cacheFieldList.item(cfi);
                            String keyString = ((Element) cf).getAttribute("name");

                            NodeList sharedItems = ((Element) cf).getElementsByTagName("sharedItems").item(0).getChildNodes();

                            List<String> values = null;

                            if (sharedItems.getLength() != 0) {
                                values = new ArrayList<>();

                                for (int sii = 0, sil = sharedItems.getLength(); sii < sil; sii++) {
                                    Node si = sharedItems.item(sii);
                                    values.add(((Element) si).getAttribute("v"));
                                }
                            }

                            definitionMap.put(keyString, values);
                        }
                    }
                }

                Document pivotCacheRecords = builder.parse(path + "/pivotCacheRecords" + (i + 1) + ".xml");
                Element pivotCacheRecordsElement = pivotCacheRecords.getDocumentElement();

                NodeList rTags = pivotCacheRecordsElement.getElementsByTagName("r");

                String[] header = definitionMap.keySet().toArray(new String[0]);

                System.out.println(header);

                String[][] result = new String[rTags.getLength()][definitionMap.size()];

                for (int ri = 0, rl = rTags.getLength(); ri < rl; ri++) {
                    Node item = rTags.item(ri);
                    NodeList columns = item.getChildNodes();


                    for (int colidx = 0, collen = columns.getLength(); colidx < collen; colidx++) {
                        Node col = columns.item(colidx);

                        String value = ((Element) col).getAttribute("v");

                        if("x".equals(col.getNodeName())){
                            value = definitionMap.get(header[colidx]).get(Integer.parseInt(value));
                        }

                        result[ri][colidx] = value;
                    }
                }

                reaultData[i] = result;
            }

            this.data = reaultData;
        } catch (ParserConfigurationException | IOException | SAXException e) {
            e.printStackTrace();
        }
    }

    public String[][] getData(int index) {
        return (String[][]) this.data[index];
    }
}
