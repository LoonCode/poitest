

import org.apache.commons.lang.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.ibatis.builder.xml.dynamic.ForEachSqlNode;
import org.apache.ibatis.executor.ExecutorException;
import org.apache.ibatis.executor.statement.BaseStatementHandler;
import org.apache.ibatis.executor.statement.RoutingStatementHandler;
import org.apache.ibatis.executor.statement.StatementHandler;
import org.apache.ibatis.mapping.BoundSql;
import org.apache.ibatis.mapping.MappedStatement;
import org.apache.ibatis.mapping.ParameterMapping;
import org.apache.ibatis.plugin.*;
import org.apache.ibatis.reflection.MetaObject;
import org.apache.ibatis.reflection.property.PropertyTokenizer;
import org.apache.ibatis.session.Configuration;
import org.apache.ibatis.session.RowBounds;
import org.apache.ibatis.type.TypeHandler;
import org.apache.ibatis.type.TypeHandlerRegistry;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.Map;
import java.util.Properties;

@Intercepts({@Signature(type = StatementHandler.class, method = "prepare", args = {Connection.class})})
public final class RecordsetPlugin implements Interceptor {

    private final static Log log = LogFactory.getLog(RecordsetPlugin.class);

    private Properties properties;

    public Object intercept(Invocation invocation) throws Throwable {

        if (!(invocation.getTarget() instanceof RoutingStatementHandler)) {
            return invocation.proceed();
        }

        RoutingStatementHandler statementHandler = (RoutingStatementHandler) invocation.getTarget();
        BoundSql boundSql = statementHandler.getBoundSql();

        // 判断是否含有分页参数，如果没有则不是分页查询
        QueryInfo queryInfo = null;
        Object parameterObject = boundSql.getParameterObject();
        if (parameterObject instanceof QueryInfo) {
            queryInfo = (QueryInfo) parameterObject;
        } else if (parameterObject instanceof Map) {
            for (Map.Entry<String, Object> e : ((Map<String, Object>) parameterObject).entrySet()) {
                if (e.getValue() instanceof QueryInfo) {
                    queryInfo = (QueryInfo) e.getValue();
                    break;
                }
            }
        }

        if (queryInfo == null || queryInfo.getRecordsetSize() == 0) {
            return invocation.proceed();
        }

        long total = this.getTotalCount(invocation);
        queryInfo.setTotalCount(total);

        MetaObject metaStatementHandler = MetaObject.forObject(statementHandler);
        RowBounds rowBounds = new RowBounds(queryInfo.getRecordsetIndex() * queryInfo.getRecordsetSize() + 1, queryInfo.getRecordsetSize());
        Configuration configuration = (Configuration) metaStatementHandler.getValue("delegate.configuration");
        Dialect.Type dialetType = getDialetType();
        if (dialetType == null) {
            throw new RuntimeException("the value of the dialect property in configuration.xml is not defined : " + configuration.getVariables().getProperty("dialect"));
        }
        Dialect dialect = null;
        switch (dialetType) {
            case MYSQL:
                dialect = new MySql5Dialect();
                break;
            case SQLSERVER:
                dialect = new SQLServerDialect();
                break;
            case ORACLE:
                dialect = new OracleDialect();
                break;
        }

        String originalSql = (String) metaStatementHandler.getValue("delegate.boundSql.sql");
        metaStatementHandler.setValue("delegate.boundSql.sql", dialect.getLimitString(originalSql, rowBounds.getOffset(), rowBounds.getLimit()));
        metaStatementHandler.setValue("delegate.rowBounds.offset", RowBounds.NO_ROW_OFFSET);
        metaStatementHandler.setValue("delegate.rowBounds.limit", RowBounds.NO_ROW_LIMIT);
        if (log.isDebugEnabled()) {
            log.debug("生成分页SQL : " + boundSql.getSql());
        }
        return invocation.proceed();

    }

    private long getTotalCount(Invocation invocation) throws Exception {

        RoutingStatementHandler statementHandler = (RoutingStatementHandler) invocation.getTarget();
        BoundSql boundSql = statementHandler.getBoundSql();
        /* 
        * 为了设置查找总数SQL的参数，必须借助MappedStatement、Configuration等这些类， 
        * 但statementHandler并没有开放相应的API，所以只好用反射来强行获取。 
        */
        BaseStatementHandler delegate = (BaseStatementHandler) ReflectionUtil.getFieldValue(statementHandler, "delegate");
        MappedStatement mappedStatement = (MappedStatement) ReflectionUtil.getFieldValue(delegate, "mappedStatement");
        Configuration configuration = mappedStatement.getConfiguration();
        TypeHandlerRegistry typeHandlerRegistry = configuration.getTypeHandlerRegistry();
        Object param = boundSql.getParameterObject();
        MetaObject metaObject = configuration.newMetaObject(param);

        long total;
        String sql = boundSql.getSql();

        // 去除order by子句
        StringBuilder sb = new StringBuilder(sql.trim());
        int orderByIndex = StringUtils.indexOfIgnoreCase(sb.toString(), "order by");
        if (orderByIndex > 0) {
            CharSequence orderby = sb.subSequence(orderByIndex, sb.length());
            sb.delete(orderByIndex, orderByIndex + orderby.length());
        }
        String countSql = "SELECT COUNT(1) FROM (" + sb.toString() + ")";
        Dialect.Type dialetType = getDialetType();
        switch (dialetType) {
            case MYSQL:               //(MYSQL、SQLSERVER要求必须添加 最后的as t)
            case SQLSERVER:
                countSql += " as t";
                break;
        }
        try {
            Connection conn = (Connection) invocation.getArgs()[0];
            PreparedStatement ps = conn.prepareStatement(countSql);
            int i = 1;
            for (ParameterMapping parameterMapping : boundSql.getParameterMappings()) {
                Object value;
                String propertyName = parameterMapping.getProperty();
                PropertyTokenizer prop = new PropertyTokenizer(propertyName);
                if (typeHandlerRegistry.hasTypeHandler(param.getClass())) {
                    value = param;
                } else if (boundSql.hasAdditionalParameter(propertyName)) {
                    value = boundSql.getAdditionalParameter(propertyName);
                } else if (propertyName.startsWith(ForEachSqlNode.ITEM_PREFIX) && boundSql.hasAdditionalParameter(prop.getName())) {
                    value = boundSql.getAdditionalParameter(prop.getName());
                    if (value != null) {
                        value = configuration.newMetaObject(value).getValue(propertyName.substring(prop.getName().length()));
                    }
                } else {
                    value = metaObject.getValue(propertyName);
                }

                TypeHandler typeHandler = parameterMapping.getTypeHandler();
                if (typeHandler == null) {
                    throw new ExecutorException("There was no TypeHandler found for parameter " + propertyName + " of statement " + mappedStatement.getId());
                }
                typeHandler.setParameter(ps, i++, value, parameterMapping.getJdbcType());
            }
            ResultSet rs = ps.executeQuery();
            rs.next();
            total = rs.getLong(1);
            rs.close();
            ps.close();
        } catch (Exception e) {
            throw new RuntimeException(e.getMessage(), e);
        }
        return total;
    }

    private Dialect.Type getDialetType() {
        return Dialect.Type.valueOf(properties.getProperty("dialect"));
    }

    public Object plugin(Object o) {

        return Plugin.wrap(o, this);
    }

    public void setProperties(Properties properties) {

        this.properties = properties;
    }

    public String getLimitString(String sql, int offset, int limit) {

        sql = sql.trim();

        StringBuffer pagingSelect = new StringBuffer(sql.length() + 100);

        pagingSelect.append("select * from ( select row_.*, rownum rownum_ from ( ");

        pagingSelect.append(sql);

        pagingSelect.append(" ) row_ ) where rownum_ >= " + offset + " and rownum_ <= " + (offset + limit - 1));


        return pagingSelect.toString();
    }
}