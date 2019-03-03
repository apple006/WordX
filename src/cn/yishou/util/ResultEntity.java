package cn.yishou.util;

import java.io.Serializable;

public class ResultEntity implements Serializable {

    public int status;
    public String error = "";
    public String message = "";
    public Object data;

    public static enum StatusCode {
        OK
    }

    public ResultEntity() {
        this.setStatus(StatusCode.OK);
    }

    public int getStatus() {
        return status;
    }

    public void setStatus(StatusCode code) {
        this.status = code.ordinal();
        this.error = code.name();
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public Object getData() {
        return data;
    }

    public void setData(Object data) {
        this.data = data;
    }

    public String getError() {
        return error;
    }

    public void setError(String error) {
        this.error = error;
    }
}
