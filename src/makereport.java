import java.io.File;  
import java.math.BigInteger;

import org.docx4j.UnitsOfMeasurement;
import org.docx4j.jaxb.Context;
import org.docx4j.model.properties.table.tc.AbstractTcProperty;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;  
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.Body;
import org.docx4j.wml.BooleanDefaultTrue;
import org.docx4j.wml.CTBorder;
import org.docx4j.wml.CTShd;
import org.docx4j.wml.HeaderReference;
import org.docx4j.wml.SectPr;
import org.w3c.dom.css.CSSPrimitiveValue;
import org.w3c.dom.css.CSSValue;

public class makereport {
	final static int c_tr=12;
	final static int c_tc=6;
	final static int c_p=6;
	final static int c_r=6;
	final static int c_t=6;
	final static int c_tcpr=6;
	final static int c_rpr=6;
	final static int c_color=6;
	final static int c_shd=6;
	final static int c_hpsmeasure=6;
	
	static WordprocessingMLPackage wordMLPackage;
	static org.docx4j.wml.ObjectFactory factory;
	
	static org.docx4j.wml.Tbl tbl;
	static org.docx4j.wml.Tr tr[]=new org.docx4j.wml.Tr[c_tr];
	static org.docx4j.wml.Tc tc[]=new org.docx4j.wml.Tc[c_tc];
	static org.docx4j.wml.P p[]=new org.docx4j.wml.P[c_p];
	static org.docx4j.wml.R r[]=new org.docx4j.wml.R[c_r];
	static org.docx4j.wml.Text t[]=new org.docx4j.wml.Text[c_t];
	
	static org.docx4j.wml.TblPr tblpr;
	static org.docx4j.wml.TcPr tcpr[]=new org.docx4j.wml.TcPr[c_tcpr];
	static org.docx4j.wml.RPr rpr[]=new org.docx4j.wml.RPr[c_rpr];
	
	static org.docx4j.wml.TblBorders tblborders;
	static org.docx4j.wml.CTBorder ctborder;
	static org.docx4j.wml.Color color[]=new org.docx4j.wml.Color[c_color];
	static org.docx4j.wml.CTShd shd[]=new org.docx4j.wml.CTShd[c_shd];
	static org.docx4j.wml.HpsMeasure hpsmeasure[]=new org.docx4j.wml.HpsMeasure[c_hpsmeasure];
	
	static org.docx4j.wml.BooleanDefaultTrue bdt=new org.docx4j.wml.BooleanDefaultTrue();
	static org.docx4j.wml.BooleanDefaultFalse bdf=new org.docx4j.wml.BooleanDefaultFalse();
	
    public static void main(String[] args) throws Exception {  
        System.out.println("begin..");  
        
        //create word file and object
        initalize();
        
        //program core
        header_factory_create();
        factory_create();
        title();
        factory_create();
        unsolved();
        factory_create();
        profile();
        factory_create();
        content();
           
        //save word file  
        wordMLPackage.save(new java.io.File(System.getProperty("user.dir") + "/aaa.docx") );  
    
        System.out.println(".. done!");  
    }
    
    
    protected static void initalize() throws InvalidFormatException{
        //create word file
        wordMLPackage=WordprocessingMLPackage.createPackage();  
        //MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
        //createHeaderPart(wordMLPackage);
        
        //open old data
        //WordprocessingMLPackage wordMLPackage=WordprocessingMLPackage.load(new java.io.File(inputfilepath));
        //MainDocumentPart documentPart=wordMLPackage.getMainDocumentPart();
        //org.docx4j.wml.Document wmlDocumentEl=(org.docx4j.wml.Document)documentPart.getJaxbElement();
        //Body body=wmlDocumentEl.getBody();
        
        factory=new org.docx4j.wml.ObjectFactory();
    }
    protected static void factory_create(){
    	int i;
    	
    	tbl=factory.createTbl();
    	for(i=0;i<c_tr;i++){
    		tr[i]=factory.createTr();
    	}
    	for(i=0;i<c_tc;i++){
    		tc[i]=factory.createTc();
    	}
    	for(i=0;i<c_p;i++){
    		p[i]=factory.createP();
    	}
    	for(i=0;i<c_r;i++){
    		r[i]=factory.createR();
    	}
    	for(i=0;i<c_t;i++){
    		t[i]=factory.createText();
    	}
        
    	tblpr=factory.createTblPr();
    	for(i=0;i<c_tcpr;i++){
    		tcpr[i]=factory.createTcPr();
    	}
    	for(i=0;i<c_rpr;i++){
    		rpr[i]=factory.createRPr();
    	}
        
    	tblborders=factory.createTblBorders();
    	ctborder=factory.createCTBorder();
    	for(i=0;i<c_color;i++){
    		color[i]=factory.createColor();
    	}
    	for(i=0;i<c_shd;i++){
    		shd[i]=factory.createCTShd();
    	}
    	for(i=0;i<c_hpsmeasure;i++){
    		hpsmeasure[i]=factory.createHpsMeasure();
    	}
    }
    
    
    protected static void title(){
    	//row 1, col 1
    	hpsmeasure[0].setVal(new java.math.BigInteger("20"));
        rpr[1].setSz(hpsmeasure[0]);
    	color[0].setVal("00FF00");
        rpr[0].setColor(color[0]);
        r[0].setRPr(rpr[0]);        
        
    	t[0].setValue("列印人員:&name&(&nowdate$)");
    	r[0].getContent().add(t[0]);
    	p[0].getContent().add(r[0]);
    	tc[0].getContent().add(p[0]);
    	tr[0].getContent().add(tc[0]);
    	
    	//row 1, col 2
    	hpsmeasure[1].setVal(new java.math.BigInteger("30"));
        rpr[1].setSz(hpsmeasure[1]);
    	rpr[1].setB(bdt);
        r[1].setRPr(rpr[1]);
        
    	t[1].setValue("國泰敦南健檢中心");
    	r[1].getContent().add(t[1]);
    	p[1].getContent().add(r[1]);
    	tc[1].getContent().add(p[1]);
    	tr[0].getContent().add(tc[1]);
    	
    	//row 1, col 3
    	hpsmeasure[2].setVal(new java.math.BigInteger("40"));
        rpr[2].setSz(hpsmeasure[2]);
    	rpr[2].setB(bdt);
        r[2].setRPr(rpr[2]);
        
        t[2].setValue("*$chartno$*");
    	r[2].getContent().add(t[2]);
    	p[2].getContent().add(r[2]);
    	tc[2].getContent().add(p[2]);
    	tr[0].getContent().add(tc[2]);
    	
    	//row 2, col 1
    	t[3].setValue("");
    	r[3].getContent().add(t[3]);
    	p[3].getContent().add(r[3]);
    	tc[3].getContent().add(p[3]);
    	tr[1].getContent().add(tc[3]);
    	
    	//row 2, col 2
    	hpsmeasure[4].setVal(new java.math.BigInteger("30"));
        rpr[1].setSz(hpsmeasure[4]);
    	rpr[4].setB(bdt);
        r[4].setRPr(rpr[4]);
        
    	t[4].setValue("敦南健檢健檢報告(院內)");
    	r[4].getContent().add(t[4]);
    	p[4].getContent().add(r[4]);
    	tc[4].getContent().add(p[4]);
    	tr[1].getContent().add(tc[4]);
    	
    	//row 2, col 3
    	t[5].setValue("");
    	r[5].getContent().add(t[5]);
    	p[5].getContent().add(r[5]);
    	tc[5].getContent().add(p[5]);
    	tr[1].getContent().add(tc[5]);
    	
    	
        tbl.getContent().add(tr[0]);
        tbl.getContent().add(tr[1]);
        wordMLPackage.getMainDocumentPart().addObject(tbl); 
    }
    
    protected static void unsolved(){
    	//row 1,col 1
    	t[0].setValue("");
    	r[0].getContent().add(t[0]);
    	p[0].getContent().add(r[0]);
    	tc[0].getContent().add(p[0]);
    	tr[0].getContent().add(tc[0]);
    	
    	
    	tbl.getContent().add(tr[0]);
    	wordMLPackage.getMainDocumentPart().addObject(tbl);
    }
    
    protected static void profile(){
    	//row 1, col 1
    	color[0].setVal("732303");
        rpr[0].setColor(color[0]);
        r[0].setRPr(rpr[0]);      
    	shd[0].setFill(UnitsOfMeasurement.rgbTripleToHex(230, 230, 230));
        tcpr[0].setShd(shd[0]);
        tc[0].setTcPr(tcpr[0]);
        
    	t[0].setValue("個人資料(PERSONAL INFORMATION)");
    	r[0].getContent().add(t[0]);
    	p[0].getContent().add(r[0]);
    	tc[0].getContent().add(p[0]);
    	tr[0].getContent().add(tc[0]);
    	
        //row 2, col 1
    	t[1].setValue("姓名(Name)");
    	r[1].getContent().add(t[1]);
    	p[1].getContent().add(r[1]);
    	tc[1].getContent().add(p[1]);
    	tr[1].getContent().add(tc[1]);
    	
    	//row 2, col 2
    	shd[2].setFill(UnitsOfMeasurement.rgbTripleToHex(230, 230, 230));
        tcpr[2].setShd(shd[2]);
        tc[2].setTcPr(tcpr[2]);
        
    	t[2].setValue("$chartname$");
    	r[2].getContent().add(t[2]);
    	p[2].getContent().add(r[2]);
    	tc[2].getContent().add(p[2]);
    	tr[1].getContent().add(tc[2]);

    	
    	tbl.getContent().add(tr[0]);
    	tbl.getContent().add(tr[1]);
    	wordMLPackage.getMainDocumentPart().addObject(tbl);
    }
    
    protected static void content(){
    	//row 1, col 1
    	color[0].setVal("732303");
        rpr[0].setColor(color[0]);
        r[0].setRPr(rpr[0]);      
    	shd[0].setFill(UnitsOfMeasurement.rgbTripleToHex(230, 230, 230));
        tcpr[0].setShd(shd[0]);
        tc[0].setTcPr(tcpr[0]);

    	t[0].setValue("目錄(CONTENTS)");
    	r[0].getContent().add(t[0]);
    	p[0].getContent().add(r[0]);
    	tc[0].getContent().add(p[0]);
    	tr[0].getContent().add(tc[0]);
        
    	//row 2,col 1
    	t[1].setValue("");
    	r[1].getContent().add(t[1]);
    	p[1].getContent().add(r[1]);
    	tc[1].getContent().add(p[1]);
    	tr[1].getContent().add(tc[1]);
        
    	tbl.getContent().add(tr[0]);
    	tbl.getContent().add(tr[1]);
    	wordMLPackage.getMainDocumentPart().addObject(tbl);
    }
    
    
    protected static void header_factory_create() throws InvalidFormatException{
        SectPr sectPr = factory.createSectPr();
        HeaderReference headerReference = factory.createHeaderReference();
        
        HeaderPart headerPart=new HeaderPart();
        Relationship relationship=wordMLPackage.getMainDocumentPart().addTargetPart(headerPart);
        //headerReference.setId(relationship.getId());
        //headerReference.setType(HdrFtrRef.DEFAULT);
        sectPr.getEGHdrFtrReferences().add(headerReference);// add header or
       
        wordMLPackage.getMainDocumentPart().addObject(sectPr);
    }
    
/*
        private static ObjectFactory objectFactory = new ObjectFactory();


        public static void createHeaderPart(
              WordprocessingMLPackage wordprocessingMLPackage)
              throws InvalidFormatException {
           HeaderPart headerPart = new HeaderPart();
            headerPart.setJaxbElement(getHdr());
           Relationship relationship = wordprocessingMLPackage.getMainDocumentPart()
                 .addTargetPart(headerPart);

           SectPr sectPr = objectFactory.createSectPr();

           HeaderReference headerReference = objectFactory.createHeaderReference();
           headerReference.setId(relationship.getId());
           headerReference.setType(HdrFtrRef.DEFAULT);
           sectPr.getEGHdrFtrReferences().add(headerReference);// add header or
          
           wordprocessingMLPackage.getMainDocumentPart().addObject(sectPr);


        }

        public static Hdr getHdr() {

           Hdr hdr = objectFactory.createHdr();

           hdr.getEGBlockLevelElts().add(getP());
           return hdr;

        }

        public static P getP() {
           P headerP = objectFactory.createP();
           R run1 = objectFactory.createR();
           Text text = objectFactory.createText();
           text.setValue("123head123");
           run1.getRunContent().add(text);
           headerP.getParagraphContent().add(run1);
           return headerP;
        }*/
        
}  


//set border line
/*ctborder.setSz(new java.math.BigInteger("10"));
tblborders.setTop(ctborder);
tblborders.setBottom(ctborder);
tblborders.setLeft(ctborder);
tblborders.setRight(ctborder);
tblpr.setTblBorders(tblborders);
tbl.setTblPr(tblpr);*/


//add by string
//wordMLPackage.getMainDocumentPart().addObject(org.docx4j.XmlUtils.unmarshalString(str));
