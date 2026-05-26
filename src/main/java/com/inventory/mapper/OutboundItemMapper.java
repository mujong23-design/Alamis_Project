package com.inventory.mapper;

import com.inventory.domain.OutboundItem;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.List;

@Mapper
public interface OutboundItemMapper {
    int insert(OutboundItem item);
    List<OutboundItem> findByOutboundId(@Param("outboundId") String outboundId);
}
