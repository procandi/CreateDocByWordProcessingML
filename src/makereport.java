public class makereport {
	public static void main(String[] args) throws Exception {
    	clsWordProcessingML wpml=new clsWordProcessingML();
    	
    	
        System.out.println("begin..");  
        
        //create word file and object
        wpml.initalize();
        
        //program core
        wpml.header_factory_create();
        wpml.header();
        wpml.footer_factory_create();
        wpml.footer();
        wpml.factory_create();
        wpml.title();
        wpml.factory_create();
        wpml.unsolved();
        wpml.factory_create();
        wpml.profile();
        wpml.factory_create();
        wpml.content();
        
        //save word file  
        wpml.wordMLPackage.save(new java.io.File(System.getProperty("user.dir") + "/aaa.docx") );  
    
        System.out.println(".. done!");  
    }        
}  