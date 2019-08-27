package com.gzmut.office.enums;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 自定义检查结果结构体
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-25
 */
@Data
@Accessors(chain = true)
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
    private StringBuffer message;

    public CheckInfoResult() {
        this.status = CheckStatus.CORRECT;
        this.score = 0.0f;
        this.message = new StringBuffer();
    }

    public CheckInfoResult(CheckStatus status, float score, String message) {
        this.status = status;
        this.score = score;
        this.message = new StringBuffer(message);
    }

    public void appendMessage(String message) {
        this.message.append(message);
    }

    private StringBuffer getBufferMessage() {
        return this.message;
    }

    public String getMessage() {
        return String.valueOf(message);
    }

    /**
     * 快速生成检查正确结果信息类
     *
     * @param score   所得分数
     * @param message 检查消息
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public static CheckInfoResult correct(float score, String message) {
        return new CheckInfoResult(CheckStatus.CORRECT, score, message + "——结果：" + CheckStatus.CORRECT.getDesc() + "——分值：" + score + System.lineSeparator());
    }

    /**
     * 快速生成检查错误结果信息类
     *
     * @param score   所得分数
     * @param message 检查消息
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public static CheckInfoResult wrong(float score, String message) {
        return new CheckInfoResult(CheckStatus.WRONG, score, message + "——结果：" + CheckStatus.WRONG.getDesc() + "——分值：" + score + System.lineSeparator());
    }

    /**
     * 快速生成检查异常结果信息类
     *
     * @param score   所得分数
     * @param message 检查消息
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public static CheckInfoResult exception(float score, String message) {
        return new CheckInfoResult(CheckStatus.EXCEPTION, score, message + "——结果：" + CheckStatus.EXCEPTION.getDesc() + "——分值：" + score + System.lineSeparator());
    }

    /**
     * 根据消息及分数追加消息记录并返回结果消息类
     *
     * @param score   所得分数
     * @param message 追加检查消息
     * @return com.gzmut.office.enums.CheckInfoResult
     */
    public CheckInfoResult appendCheckInfoResult(CheckStatus checkStatus, float score, String message) {
        this.setScore(this.getScore() + score);
        if(message != null){
            this.getBufferMessage().append(message).append("——结果：").append(checkStatus.getDesc()).append("——分值：").append(score).append(System.lineSeparator());
        }
        if (checkStatus != CheckStatus.CORRECT) {
            this.setStatus(checkStatus);
        }
        return this;
    }
}

