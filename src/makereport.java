import java.io.File;  

import org.docx4j.UnitsOfMeasurement;
import org.docx4j.jaxb.Context;
import org.docx4j.model.properties.table.tc.AbstractTcProperty;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;  
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Body;
import org.docx4j.wml.CTShd;
import org.w3c.dom.css.CSSPrimitiveValue;
import org.w3c.dom.css.CSSValue;

public class makereport {
    public static void main(String[] args) throws Exception {  
        System.out.println("begin..");  
        
        String inputfilepath;
        inputfilepath=System.getProperty("user.dir") + "/test.docx";
        
        //create word file
        WordprocessingMLPackage wordMLPackage=WordprocessingMLPackage.createPackage();  
        //WordprocessingMLPackage wordMLPackage=WordprocessingMLPackage.load(new java.io.File(inputfilepath));
        //MainDocumentPart documentPart=wordMLPackage.getMainDocumentPart();
        //org.docx4j.wml.Document wmlDocumentEl=(org.docx4j.wml.Document)documentPart.getJaxbElement();
        //Body body=wmlDocumentEl.getBody();
        
        org.docx4j.wml.ObjectFactory factory=new org.docx4j.wml.ObjectFactory();
        
        org.docx4j.wml.Tbl tbl=factory.createTbl();
        org.docx4j.wml.Tr tr=factory.createTr();
        org.docx4j.wml.Tc tc=factory.createTc();
        org.docx4j.wml.P p=factory.createP();
        org.docx4j.wml.R r=factory.createR();
        org.docx4j.wml.Text t=factory.createText();
        
        org.docx4j.wml.RPr rpr=factory.createRPr();
        org.docx4j.wml.Color color=factory.createColor();
        color.setVal("FF0000");
        rpr.setColor(color);
        r.setRPr(rpr);
        
        org.docx4j.wml.TcPr tcpr=factory.createTcPr();
        CTShd shd=factory.createCTShd();
    	short ignored=1;
    	/*
    	CSSValue value;
    	value="";
    	CSSPrimitiveValue cssPrimitiveValue=(CSSPrimitiveValue)value;
    	float fRed=cssPrimitiveValue.getRGBColorValue().getRed().getFloatValue(ignored);
    	float fGreen=cssPrimitiveValue.getRGBColorValue().getGreen().getFloatValue(ignored);
    	float fBlue=cssPrimitiveValue.getRGBColorValue().getBlue().getFloatValue(ignored);
    	shd.setFill(UnitsOfMeasurement.rgbTripleToHex(fRed, fGreen, fBlue));
    	*/
    	shd.setFill(UnitsOfMeasurement.rgbTripleToHex(55, 55, 55));
        tcpr.setShd(shd);
        tc.setTcPr(tcpr);
        
        t.setValue("text");
        
        r.getContent().add(t);
        p.getContent().add(r);
        tc.getContent().add(p);
        tr.getContent().add(tc);
        tbl.getContent().add(tr);
        wordMLPackage.getMainDocumentPart().addObject(tbl);
        //wordMLPackage.getMainDocumentPart().addObject(org.docx4j.XmlUtils.unmarshalString(str));  
        
        //save word file  
        wordMLPackage.save(new java.io.File(System.getProperty("user.dir") + "/aaa.docx") );  
    
        System.out.println(".. done!");  
    }
}  