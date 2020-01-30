package cn.hillwind.tools.excel;

public class Cell {
    private Object value;
    private CellType type;

    public Cell() {
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public CellType getType() {
        return type;
    }

    public void setType(CellType type) {
        this.type = type;
    }
}
