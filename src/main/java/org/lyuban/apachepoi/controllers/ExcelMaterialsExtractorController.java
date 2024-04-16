package org.lyuban.apachepoi.controllers;

import org.lyuban.apachepoi.model.MaterialPosition;
import org.lyuban.apachepoi.servises.ExcelMaterialsExtractorService;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/get-project-materials")
public class ExcelMaterialsExtractorController {
    private final ExcelMaterialsExtractorService emes;

    public ExcelMaterialsExtractorController(ExcelMaterialsExtractorService emes) {
        this.emes = emes;
    }
    public ResponseEntity<MaterialPosition> getProjectAllMaterials(){
        emes.
    }
}
