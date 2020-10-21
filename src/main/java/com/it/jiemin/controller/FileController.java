package com.it.jiemin.controller;

import com.it.jiemin.common.EncodeUtils;
import com.it.jiemin.common.ExcOperator;
import com.it.jiemin.common.Response;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.IOException;

/**
 * @author JieminZhou
 * @version 1.0
 * @date 2020/10/20 16:36
 */
@RestController
@RequestMapping("/")
public class FileController {
    /**
     * 合并多个Excel中的内容
     * @param files 待合并的Excel
     * @param request 请求
     * @return 合并后的Excel
     */
    @PostMapping(value = "/mergeFiles")
    public ResponseEntity<ByteArrayResource> mergeFiles(@RequestParam(value = "files") MultipartFile[] files,
                                                    HttpServletRequest request) {
        try {
            InputStreamResource[] fileInputAsResources = new InputStreamResource[files.length];
            for(int i = 0; i < files.length; i++) {
                MultipartFile file = files[i];
                fileInputAsResources[i] = new InputStreamResource(file.getInputStream()){
                    @Override
                    public String getFilename() {
                        return file.getOriginalFilename();
                    }

                    @Override
                    public long contentLength() throws IOException {
                        return file.getSize();
                    }
                };
            }
            byte[] bytes = ExcOperator.merge(fileInputAsResources);
            ByteArrayResource br = new ByteArrayResource(bytes);
            return ResponseEntity.ok()
                    .contentType(MediaType.parseMediaType("application/octet-stream"))
                    .contentLength(bytes.length)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\""
                            + EncodeUtils.getFilename(request, fileInputAsResources[0].getFilename()) + "\"")
                    .body(br);
        } catch (Exception ex) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(new ByteArrayResource(ex.getMessage().getBytes()));
        }
    }
}
