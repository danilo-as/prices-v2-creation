-- auto-generated definition
create table grades
(
    id                       uuid      default uuid_generate_v4() not null
        primary key,
    internal_code            varchar(50)                          not null
        unique,
    full_name                varchar(255)                         not null,
    short_name               varchar(200),
    specification            text,
    market_id                uuid                                 not null
        references primary_data.market
            on update restrict on delete restrict,
    market_name              varchar(100)                         not null,
    price_category_id        uuid                                 not null
        constraint grades_price_category_id_fk
            references price_category,
    price_category_name      varchar(100)                         not null,
    product_id               uuid                                 not null
        constraint grades_product_id_fk
            references product,
    product_name             varchar(100)                         not null,
    product_alias            varchar(100),
    capacity_id              uuid
        constraint grades_capacity_id_fk
            references capacity,
    capacity_name            varchar(100),
    thickness_id             uuid
        constraint grades_thickness_id_fk
            references thickness,
    thickness_name           varchar(100),
    cell_format_id           uuid
        constraint grades_cell_format_id_fk
            references cell_format,
    cell_format_name         varchar(100),
    country_id               uuid
        constraint grades_country_id_fk
            references country,
    country_name             varchar(100),
    mesh_size_id             uuid
        constraint grades_mesh_size_id_fk
            references mesh_size,
    mesh_size_name           varchar(100),
    purity_id                uuid
        constraint grades_purity_id_fk
            references purity,
    purity_name              varchar(100),
    incoterm_id              uuid
        constraint grades_incoterm_id_fk
            references incoterm,
    incoterm_name            varchar(100),
    region_id                uuid
        constraint grades_region_id_fk
            references region,
    region_name              varchar(100),
    service_id               uuid
        constraint grades_service_id_fk
            references service,
    service_name             varchar(100),
    sub_product_id           uuid
        constraint grades_sub_product_id_fk
            references sub_product,
    sub_product_name         varchar(100),
    feedstock_id             uuid
        constraint grades_feedstock_id_fk
            references feedstock,
    feedstock_name           varchar(100),
    trade_type_id            uuid
        constraint grades_trade_type_id_fk
            references trade_type,
    trade_type_name          varchar(100),
    grade                    varchar(50),
    price_type_id            uuid
        constraint grades_price_type_id_fk
            references price_type,
    price_type_name          varchar(100),
    unit_of_measure_id       uuid                                 not null
        constraint grades_unit_of_measure_id_fk
            references unit_of_measure,
    unit_of_measure_name     varchar(100)                         not null,
    additional_configuration jsonb,
    frequency_id             uuid                                 not null
        constraint grades_frequencies_id_fk
            references ??? (),
    frequency_name           varchar(100)                         not null,
    default_currency_id      uuid
        constraint grades_default_currency_id_fk
            references primary_data.currency,
    "order"                  integer                              not null,
    assessment_launched_at   date,
    last_assessed_at         date,
    is_sustainable           boolean                              not null,
    is_iosco_assured         boolean   default false              not null,
    is_public                boolean   default false              not null,
    has_only_price_mid       boolean   default false              not null,
    is_active                boolean                              not null,
    is_spot                  boolean   default true,
    is_price_grade           boolean   default true,
    created_by               uuid,
    created_at               timestamp default now(),
    updated_by               uuid,
    updated_at               timestamp,
    unique_sha256            text generated always as (encode(digest(
                                                                      (((((((((((((((((((((((((((((((((((((((frequency_id)::text || '-'::text) || (market_id)::text) ||
                                                                                                          '-'::text) ||
                                                                                                         (price_category_id)::text) ||
                                                                                                        '-'::text) ||
                                                                                                       (product_id)::text) ||
                                                                                                      '-'::text) ||
                                                                                                     COALESCE((capacity_id)::text, ''::text)) ||
                                                                                                    '-'::text) ||
                                                                                                   COALESCE((thickness_id)::text, ''::text)) ||
                                                                                                  '-'::text) ||
                                                                                                 COALESCE((cell_format_id)::text, ''::text)) ||
                                                                                                '-'::text) ||
                                                                                               COALESCE((country_id)::text, ''::text)) ||
                                                                                              '-'::text) ||
                                                                                             COALESCE((mesh_size_id)::text, ''::text)) ||
                                                                                            '-'::text) ||
                                                                                           COALESCE((purity_id)::text, ''::text)) ||
                                                                                          '-'::text) ||
                                                                                         COALESCE((incoterm_id)::text, ''::text)) ||
                                                                                        '-'::text) ||
                                                                                       COALESCE((region_id)::text, ''::text)) ||
                                                                                      '-'::text) ||
                                                                                     COALESCE((service_id)::text, ''::text)) ||
                                                                                    '-'::text) ||
                                                                                   COALESCE((sub_product_id)::text, ''::text)) ||
                                                                                  '-'::text) ||
                                                                                 COALESCE((feedstock_id)::text, ''::text)) ||
                                                                                '-'::text) ||
                                                                               COALESCE((trade_type_id)::text, ''::text)) ||
                                                                              '-'::text) ||
                                                                             COALESCE((grade)::text, ''::text)) ||
                                                                            '-'::text) ||
                                                                           COALESCE((unit_of_measure_id)::text, ''::text)) ||
                                                                          '-'::text) ||
                                                                         COALESCE((price_type_id)::text, ''::text)) ||
                                                                        '-'::text) || (is_sustainable)::text),
                                                                      'sha256'::text), 'hex'::text)) stored
        constraint grade_unique_pk
            unique,
    sulphur_id               uuid
                                                                  references sulphur
                                                                      on update cascade on delete set null,
    sulphur_name             varchar(100),
    constraint grades_name_market_id_pk
        unique (full_name, market_id)
);

alter table grades
    owner to postgres;

