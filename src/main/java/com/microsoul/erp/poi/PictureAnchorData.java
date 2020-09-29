package com.microsoul.erp.poi;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;

/**
 * 微妞分布式平台-公用工具包
 * @author 广州加叁信息科技有限公司 (tiger@microsoul.com)
 * @version V1.0.0
 */
public class PictureAnchorData {
    private byte[] data;
    private HSSFClientAnchor anchor;

    public PictureAnchorData(HSSFClientAnchor anchor, byte[] data) {
        this.anchor = anchor;
        this.data = data;
    }

    public byte[] getData() {
        return data;
    }

    public void setData(byte[] data) {
        this.data = data;
    }

    public HSSFClientAnchor getAnchor() {
        return anchor;
    }

    public void setAnchor(HSSFClientAnchor anchor) {
        this.anchor = anchor;
    }
}
