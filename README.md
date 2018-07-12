# fact-excel

## **当前版本**

1.0.0-SNAPSHOT

## **Maven依赖**

```xml

<dependency>

<groupId>com.woter.fact</groupId>

<artifactId>fact-excel</artifactId>

<version>1.0.0-SNAPSHOT</version>

</dependency>

```

## **功能描述**

fact-excel 是基于POI封装常用的Excel导出功能；主要包括一下几个方面的功能，具体如下：

- [x] 开箱即用，支持WEB或者文件流方式导出；  

- [x] 内置默认配置也支持每个字段自定义配置；  

- [x] 内置常见的日期、时间字段等解析器；  

- [x] 超过最大行会自动创建下一个sheet；  

## **升级日志** 
   暂无


## **常用功能代码演示**

### ** 示例：导出用户列表 **
```java

public class User {
    
    private String name;
    
    private Integer age;
    
    private Date brithday;

    public String getName() {
        return name;
    }
    
    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }
    
    public void setAge(Integer age) {
        this.age = age;
    }

    public Date getBrithday() {
        return brithday;
    }

    public void setBrithday(Date brithday) {
        this.brithday = brithday;
    }
 
}

@Controller
@RequestMapping("/demo")
public void DemoController{
    
    @RequestMapping("/export")
    public void exportExcel(HttpServletResponse response){
        
        List<User> userList = userService.selectSignaturePagingListByParam(userSearchParam);
        
        List<ExcelField<?>> fieldList = Lists.newArrayList();
        fieldList.add(new ExcelField<String>("姓名", "name"));
        fieldList.add(new ExcelField<Integer>("年龄", "age"));
        fieldList.add(new ExcelField<Object>("生日",FieldTranslationFacade.dateFieldTranslation("brithday")));
        try {
            ExcelOperationFacade.exportExcel("用户列表", fieldList, (List<?>) userList, response, "用户列表");
        } catch (Exception e) {
            logger.error("导出用户列表异常",e);
        }
    }

}

```
