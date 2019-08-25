package com.gzmut.office.enums;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 自定义检查结果结构体
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-25
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class CheckInfoResult {
    /**
     * 检查状态
     */
    private CheckStatus status;

    /**
     * 所得分数
     */
    private float score;

    /**
     * 日志消息
     */
    private String message;

    public static CheckInfoResult correct(float score,String message){
        return new CheckInfoResult(CheckStatus.CORRECT,score,message);
    }

    public static CheckInfoResult wrong(float score,String message){
        return new CheckInfoResult(CheckStatus.WRONG,score,message);
    }

    public static CheckInfoResult exception(float score,String message){
        return new CheckInfoResult(CheckStatus.EXCEPTION,score,message);
    }

}

