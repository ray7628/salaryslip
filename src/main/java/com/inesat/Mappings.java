package com.inesat;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.List;

@Data
@Component
@ConfigurationProperties(prefix = "mapping")
public class Mappings {
    private List<ColMapping> cols;
}

@Data
class ColMapping {
    private String from;
    private String to;
    private Integer colIndex;
}
