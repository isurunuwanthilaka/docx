package com.isuru.docx;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.util.HashMap;

@RestController
@RequestMapping("v1")
public class DocxResource {

    @GetMapping("doc")
    public void getDoc() throws Exception{
        org.docx4j.wml.ObjectFactory foo = Context.getWmlObjectFactory();

        String inputfilepath = System.getProperty("user.dir") + "/templates/example.docx";

        boolean save = true;

        String outputfilepath = System.getProperty("user.dir")
                + "/templates/example-filled.docx";

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                .load(new java.io.File(inputfilepath));
        MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

        HashMap<String, String> mappings = new HashMap<String, String>();
        mappings.put("colour", "green");
        mappings.put("value", "Hundred");


        documentPart.variableReplace(mappings);


        // Save it
        if (save)

        {
            SaveToZipFile saver = new SaveToZipFile(wordMLPackage);
            saver.save(outputfilepath);
        } else

        {
            System.out.println(XmlUtils.marshaltoString(documentPart.getJaxbElement(), true,
                    true));
        }

    }

}
