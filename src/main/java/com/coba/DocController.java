package com.coba;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class DocController {

	private static final String SOURCEFILE = "/doc/skt-bilingual.doc";
    private static final String OUTPUTFILE = "/home/ahmad/Downloads/skt-bilingual-new.doc";

	@GetMapping("/tes")
	public void print() throws Exception {
		
        HWPFDocument doc = this.openDocument(SOURCEFILE);
        if (doc != null) {
            doc = this.replaceText(doc);
            this.saveDocument(doc, OUTPUTFILE);
        }
	}

    private HWPFDocument replaceText(HWPFDocument doc) {
    	
        Range r = doc.getRange();
        for (int i = 0; i < r.numSections(); ++i) {
            Section s = r.getSection(i);
            for (int j = 0; j < s.numParagraphs(); j++) {
                Paragraph p = s.getParagraph(j);
                for (int k = 0; k < p.numCharacterRuns(); k++) {
                    CharacterRun run = p.getCharacterRun(k);
                    String text = run.text();
                    if (text.contains("NOMORSKT")) {
                        run.replaceText("NOMORSKT", "SKT-001/WPJ/2020");
                    } else if (text.contains("NOMORKEP")) {
                        run.replaceText("NOMORKEP", "KEP-001/WPJ/2020");
                    } else if (text.contains("NAMAWP")) {
                        run.replaceText("NAMAWP", "Ahmad Surya Putra");
                    } else if (text.contains("NPWP")) {
                        run.replaceText("NPWP", "0100000000000");
                    } else if (text.contains("ALAMATWP")) {
                        run.replaceText("ALAMATWP", "Jalan Gatsu");
                    } else if (text.contains("EMAILWP")) {
                        run.replaceText("EMAILWP", "ahmad.putra91@gmail.com");
                    }  else if (text.contains("KATEGORIWP")) {
                        run.replaceText("KATEGORIWP", "PERSON");
                    }  else if (text.contains("TGLKEPID")) {
                        run.replaceText("TGLKEPID", "01-01-2020");
                    }  else if (text.contains("TGLKEP")) {
                        run.replaceText("TGLKEP", "2020-01-01");
                    }  else if (text.contains("TGLTERDAFTARID")) {
                        run.replaceText("TGLTERDAFTARID", "01-01-2020");
                    } else if (text.contains("TGLTERDAFTAR")) {
                        run.replaceText("TGLTERDAFTAR", "2020-01-01");
                    }   else if (text.contains("TGLPENUNJUKANID")) {
                        run.replaceText("TGLPENUNJUKANID", "01-01-2020");
                    } else if (text.contains("TGLPENUNJUKAN")) {
                        run.replaceText("TGLPENUNJUKAN", "2020-01-01");
                    }  else if (text.contains("TGLSKT")) {
                        run.replaceText("TGLSKT", "2020-01-01");
                    }  else if (text.contains("NAMAKASI")) {
                        run.replaceText("NAMAKASI", "Ahmad Surya Putra");
                    }
                }
            }
        }
        return doc;
    }

    private HWPFDocument openDocument(String file) throws Exception {
        //URL res = getClass().getClassLoader().getResource(file);
        Resource resource = new ClassPathResource(file);
        HWPFDocument document = null;
        if (resource.getFile() != null) {
            document = new HWPFDocument(new POIFSFileSystem(resource.getFile()));
        }
        return document;
    }

    private void saveDocument(HWPFDocument doc, String file) {
        try (FileOutputStream out = new FileOutputStream(file)) {
            doc.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
