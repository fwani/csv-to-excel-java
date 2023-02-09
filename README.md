# CSV to Excel

`.csv` 파일을 excel(`.xls`, `.xlsx`) 파일로 변환하는 코드

## 변경점

`0.0.1` 버전은 sheet 의 `MAX_ROW`(`default: 1,048,576`) 를 넘어가는 데이터는 버린다.

## 사용방법

### 0. 세팅 추가

- `settings.xml` 설정

```xml
<server>
    <id>github</id>
    <username>${github 유저명}</username>
    <password>${github 토큰}</password>
</server>
```

- `pom.xml` 에 설정 추가

```xml
<project>
    <dependencies>
        <dependency>
            <groupId>pe.fwani.convert</groupId>
            <artifactId>csv-to-excel</artifactId>
            <version>0.0.1-java11-SNAPSHOT</version>
        </dependency>
    </dependencies>

    <repositories>
        <repository>
            <id>github-fwani-releases</id>
            <url>https://github.com/fwani/fwani-maven-repo/raw/main/releases</url>
            <releases>
                <enabled>true</enabled>
                <updatePolicy>always</updatePolicy>
            </releases>
        </repository>
        <repository>
            <id>github-fwani-snapshots</id>
            <url>https://github.com/fwani/fwani-maven-repo/raw/main/snapshots</url>
            <snapshots>
                <enabled>true</enabled>
                <updatePolicy>always</updatePolicy>
            </snapshots>
        </repository>
    </repositories>
</project>
```

### 1. 전체 데이터 변환

```java
import pe.fwani.convert.CsvToExcel;

class Example {
    public static void main(String[] args) {
        var convertor = new CsvToExcel("origin.csv", "output.xlsx");
        convertor.setWorkbookType(WorkbookType.SXSSF);
        convertor.convert("시트1");
    }
}
```

### 2. 각 시트별 특정 컬럼 선택

```java
import pe.fwani.convert.CsvToExcel;

class Example {
    public static void main(String[] args) {
        var convertor = new CsvToExcel("origin.csv", "output.xlsx");
        convertor.setWorkbookType(WorkbookType.SXSSF);
        convertor.convert(List.of(
                new Pair<>("시트1", List.of("col1", "col2")),
                new Pair<>("시트2", List.of("col1", "col3", "col4"))
        ));
    }
}
```

## Contribute

### Setting

- `settings.xml` 설정

```xml
<server>
    <id>github</id>
    <username>${github 유저명}</username>
    <password>${github 토큰}</password>
</server>
```

### Deploy

- 배포
  - 배포 결과는 [fwani-maven-repo](https://github.com/fwani/fwani-maven-repo) 레파지토리 확인

```shell
./mvnw -Drevision=<버전> clean deploy -f pom.xml
```
