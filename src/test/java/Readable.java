import com.cemgunduz.model.annotation.ExcelMapping;
import com.cemgunduz.model.annotation.ExcelMappings;

/**
 * Created by cgunduz on 2/7/15.
 */
public class Readable {

    @ExcelMappings(list = {@ExcelMapping (coloumnNo = 0), @ExcelMapping(key = "AnotherMapping", coloumnNo = 3)})
    private Long id;

    @ExcelMappings(list = {@ExcelMapping (coloumnNo = 1), @ExcelMapping(key = "AnotherMapping", coloumnNo = 24)})
    private String name;

    @ExcelMappings(list = {@ExcelMapping (coloumnNo = 2), @ExcelMapping(key = "AnotherMapping", coloumnNo = 5)})
    private String surname;

    @ExcelMappings(list = {@ExcelMapping (coloumnNo = 3), @ExcelMapping(key = "AnotherMapping", coloumnNo = 19)})
    private Double money;

    // Getter setters

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Double getMoney() {
        return money;
    }

    public void setMoney(Double money) {
        this.money = money;
    }

    public String getSurname() {
        return surname;
    }

    public void setSurname(String surname) {
        this.surname = surname;
    }
}
