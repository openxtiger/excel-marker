package com.microsoul.erp.poi;

import org.apache.poi.ss.usermodel.ClientAnchor;

/**
 * 微妞分布式平台-公用工具包
 * @author 广州加叁信息科技有限公司 (tiger@microsoul.com)
 * @version V1.0.0
 */
public interface POIHelper {

    void createPicture(ClientAnchor anchor, int width, int height, byte[] pictureData, int insidePictureCount);

    void createPicture(int col, int row, int px1, int py1, int width, int height, byte[] pictureData, int insidePictureCount);

    void copyCell(int rowStart, int colStart, int colEnd, int dataRows, int offset, int count);

}
