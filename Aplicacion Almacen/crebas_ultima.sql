rem ============================================================
rem   Database name:  LSI_BBDD                                  
rem   DBMS name:      ORACLE Version 6.0                        
rem   Created on:     24/11/2002  21:17                         
rem ============================================================

rem ============================================================
rem   Eliminaci¢n de tablas
rem ============================================================
drop table pais;
drop table linea_pedido;
drop table producto_almacen;
drop table producto;
drop table almacen;
drop table region;
drop table proveedor;
drop table incidencias;
drop table orden_pedido;
drop table cliente;
drop table empleado;

rem ============================================================
rem   Table: REGION                                             
rem ============================================================
create table REGION
(
    CODIGO_REGION          VARCHAR(3)               not null,
    NOMBRE                 VARCHAR(20)                      ,
    constraint PK_REGION primary key (CODIGO_REGION) 
);

rem ============================================================
rem   Table: PROVEEDOR                                          
rem ============================================================
create table PROVEEDOR
(
    CODIGO                 VARCHAR(3)               not null,
    NIF_CIF                VARCHAR(9)               not null,
    NOMBRE                 VARCHAR(20)                      ,
    TELEFONO               VARCHAR(12)                      ,
    FAX                    VARCHAR(12)                      ,
    EMAIL                  VARCHAR(60)                      ,
    DIRECCION              VARCHAR(40)                      ,
    NUMERO                 NUMBER(4)                     ,
    PUERTA                 NUMBER(4)                     ,
    LOCALIDAD              VARCHAR(20)                      ,
    PROVINCIA              VARCHAR(20)                      ,
    CP                     NUMBER(5)                     ,
    PAIS                   VARCHAR(15)                      ,
    constraint PK_PROVEEDOR primary key (CODIGO) 
);

rem ============================================================
rem   Table: EMPLEADO                                           
rem ============================================================
create table EMPLEADO
(
    NIF                    VARCHAR(9)               not null,
    LOGIN                  VARCHAR(8)               not null,
    PASSWORD               VARCHAR(8)               not null,
    NOMBRE                 VARCHAR(30)                      ,
    CARGO                  VARCHAR(20)              not null,
    TELEFONO               VARCHAR(12)                      ,
    constraint PK_EMPLEADO primary key (NIF) 
);

rem ============================================================
rem   Table: PRODUCTO                                           
rem ============================================================
create table PRODUCTO
(
    REFERENCIA             VARCHAR(5)               not null,
    CODIGO_PROVEEDOR       VARCHAR(3)               not null,
    NOMBRE			   VARCHAR(20)			         ,
    DESCRIPCION            VARCHAR(255)                     ,
    PRECIO                 NUMBER(5,2)           not null,
    MAX_RAZONABLE          NUMBER(4)                     ,
    constraint PK_PRODUCTO primary key (REFERENCIA) 
);

rem ============================================================
rem   Table: ALMACEN                                            
rem ============================================================
create table ALMACEN
(
    CODIGO_ALMACEN         VARCHAR(3)             not null,
    CODIGO_REGION          VARCHAR(3)             unique not null,
    TECNICO_ALMACEN        VARCHAR(9)             unique not null,
    NOMBRE                 VARCHAR(20)                      ,
    DIRECCION              VARCHAR(40)                      ,
    PAIS                   VARCHAR(15)                      ,
    TELEFONO               VARCHAR(12)                      ,
    FAX                    VARCHAR(12)                      ,
    EMAIL                  VARCHAR(60)                      ,
   constraint PK_ALMACEN primary key (CODIGO_ALMACEN) 
);

rem ============================================================
rem   Table: CLIENTE                                            
rem ============================================================
create table CLIENTE
(
    NIF_CIF                VARCHAR(9)        	 unique not null,
    CODIGO			   VARCHAR(5)			 not null,
    REPRESENTANTE          VARCHAR(9)                       ,
    LOGIN                  VARCHAR(8)               	   ,
    PASSWORD               VARCHAR(8)               	   ,
    NOMBRE                 VARCHAR(40)              not null,
    CALLE                  VARCHAR(30)              not null,
    NUMERO                 NUMBER(4)             ,
    PUERTA                 NUMBER(4)             ,
    LOCALIDAD              VARCHAR(20)              not null,
    PROVINCIA              VARCHAR(20)              not null,
    CP                     NUMBER(5)                     ,
    PAIS                   VARCHAR(15)              not null,
    TELEFONO               VARCHAR(12)              not null,
    FAX                    VARCHAR(12)                      ,
    EMAIL                  VARCHAR(60)                      ,
    CUENTA_BANCO           VARCHAR(25)              not null,
    RATIO_CONFIANZA        VARCHAR(9)             not null,
    ES_EMPRESA             NUMBER(1)             not null,
    PERSONA_CONTACTO       VARCHAR(30)                      ,
    TLF_PERS_CONTACTO      VARCHAR(12)                      ,
   constraint PK_CLIENTE primary key (CODIGO) 
);

rem ============================================================
rem   Table: ORDEN_PEDIDO                                       
rem ============================================================
create table ORDEN_PEDIDO
(
    CODIGO_PEDIDO          NUMBER(5)             not null,
    USUARIO_VENTAS         VARCHAR(9)               not null,
    CLIENTE                VARCHAR(5)               not null,
    CALLE_ENVIO            VARCHAR(30)              not null,
    NUMERO_ENVIO           NUMBER(4)             not null,
    PTA_ENVIO           NUMBER(4)             not null,
    LOCALIDAD_ENVIO        VARCHAR(20)              not null,
    PROVINCIA_ENVIO        VARCHAR(20)              not null,
    CP_ENVIO               NUMBER(5)             not null,
    PAIS_ENVIO             VARCHAR(15)              not null,
    FORMA_PAGO             NUMBER(1)             not null,
    FECHA_ELABORACION      DATE                  not null,
    FECHA_LLEGADA_ALMACEN  DATE                          ,
    FECHA_ATENCION         DATE                          ,
    FECHA_LISTO_ENVIO      DATE                          ,
    FECHA_SALIDA_ALMACEN   DATE                          ,
   constraint PK_ORDEN_PEDIDO primary key (CODIGO_PEDIDO) 
);

rem ============================================================
rem   Table: INCIDENCIAS                                               
rem ============================================================
create table INCIDENCIAS
(
    CODIGO_INCIDENCIA      NUMBER(5)              not null,
    CODIGO_PEDIDO          NUMBER(5)               not null,
    FECHA_INCIDENCIA       DATE                 not null,
    NIF_CREADOR            VARCHAR(9)          not null,
    CREADOR                VARCHAR(30)            not null,
    OBSERVACIONES          VARCHAR(255)          not null,    
   constraint PK_INCIDENCIAS primary key (CODIGO_INCIDENCIA) 
);

rem ============================================================
rem   Table: PAIS                                               
rem ============================================================
create table PAIS
(
    NOMBRE                 VARCHAR(15)              not null,
    CODIGO_REGION          VARCHAR(3)               not null,
   constraint PK_PAIS primary key (NOMBRE) 
);

rem ============================================================
rem   Table: PRODUCTO_ALMACEN                                   
rem ============================================================
create table PRODUCTO_ALMACEN
(
    REFERENCIA             VARCHAR(5)               not null,
    ALMACEN		         VARCHAR(3)               not null,
    STOCK                  NUMBER(6)                not null,
    STOCK_ASIGNADO         NUMBER(6)      DEFAULT 0 not null,
   constraint PK_PROD_ALMACEN primary key (REFERENCIA, ALMACEN) 
);

rem ============================================================
rem   Table: LINEA_PEDIDO                                       
rem ============================================================
create table LINEA_PEDIDO
(
    REFERENCIA             VARCHAR(5)            not null,
    CODIGO_PEDIDO          NUMBER(5)             not null,
    CANTIDAD               NUMBER(5)             not null,
    PRECIO		         NUMBER(5,2)           not null,
    CANT_ASIGNADA          NUMBER(5)             DEFAULT 0 not null,
   constraint PK_LINEA_PEDIDO primary key (REFERENCIA, CODIGO_PEDIDO) 
);

alter table PRODUCTO
    add constraint FK_PRODUCTO_REF_161_PROVEEDO foreign key  (CODIGO_PROVEEDOR)
       references PROVEEDOR (CODIGO);

alter table ALMACEN
    add constraint FK_ALMACEN_REF_97_REGION foreign key  (CODIGO_REGION)
       references REGION (CODIGO_REGION);

alter table ALMACEN
    add constraint FK_ALMACEN_REF_111_EMPLEADO foreign key  (TECNICO_ALMACEN)
       references EMPLEADO (NIF);

alter table CLIENTE
    add constraint FK_CLIENTE_REF_121_EMPLEADO foreign key  (REPRESENTANTE)
       references EMPLEADO (NIF);

alter table ORDEN_PEDIDO
    add constraint FK_ORDEN_PE_REF_131_EMPLEADO foreign key  (USUARIO_VENTAS)
       references EMPLEADO (NIF);

alter table ORDEN_PEDIDO
    add constraint FK_ORDEN_PE_REF_141_CLIENTE foreign key  (CLIENTE)
       references CLIENTE (CODIGO);

alter table PAIS
     add constraint FK_PAIS_REF_87_REGION foreign key  (CODIGO_REGION)
       references REGION (CODIGO_REGION);

alter table PRODUCTO_ALMACEN
    add constraint FK_PRODUCTO_REF_171_PRODUCTO foreign key  (REFERENCIA)
       references PRODUCTO (REFERENCIA);

alter table PRODUCTO_ALMACEN
    add constraint FK_PRODUCTO_REF_175_ALMACEN foreign key  (ALMACEN)
       references ALMACEN (CODIGO_ALMACEN);

alter table LINEA_PEDIDO
    add constraint FK_LINEA_PE_REF_193_PRODUCTO foreign key  (REFERENCIA)
       references PRODUCTO (REFERENCIA);

alter table LINEA_PEDIDO
    add constraint FK_LINEA_PE_REF_199_ORDEN_PE foreign key  (CODIGO_PEDIDO)
       references ORDEN_PEDIDO (CODIGO_PEDIDO) ON DELETE CASCADE;

alter table INCIDENCIAS
    add constraint FK_INCIDENC_REF_205_ORDEN_PE foreign key  (CODIGO_PEDIDO)
       references ORDEN_PEDIDO (CODIGO_PEDIDO) ON DELETE CASCADE;

alter table INCIDENCIAS
    add constraint FK_INCIDENC_REF_215_EMPLEADO foreign key  (NIF_CREADOR)
       references EMPLEADO (NIF);

