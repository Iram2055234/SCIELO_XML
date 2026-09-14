<?xml version='1.0' encoding='UTF-8'?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:xlink="http://www.w3.org/1999/xlink" version="1.0" exclude-result-prefixes="xlink">
    
    <xsl:output method="html" encoding="UTF-8" indent="yes" doctype-system="about:legacy-compat"/>
    
    <xsl:strip-space elements="*"/>
    
    <xsl:template match="/">
        <xsl:variable name="article-meta" select="article/front/article-meta"/>
        
        <xsl:variable name="editorial-meta" select="$article-meta/custom-meta-group/custom-meta"/>
        
        <html lang="{article/@xml:lang}">
            <head>
                <meta charset="UTF-8"/>
                
                <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
                
                <title>
                    <xsl:value-of select="$article-meta/title-group/article-title"/>
                </title>
                
                <style>
                    * {
                    box-sizing: border-box;
                    }
                    
                    html {
                    scroll-behavior: smooth;
                    }
                    
                    body {
                    margin: 0;
                    background: #eeeeee;
                    color: #202020;
                    font-family: Georgia, "Times New Roman", serif;
                    font-size: 17px;
                    line-height: 1.65;
                    }
                    
                    a {
                    color: #135f7b;
                    text-decoration-thickness: 1px;
                    text-underline-offset: 2px;
                    overflow-wrap: anywhere;
                    }
                    
                    .article-container {
                    width: min(1380px, calc(100% - 32px));
                    margin: 28px auto;
                    background: #ffffff;
                    box-shadow: 0 3px 18px rgba(0, 0, 0, 0.10);
                    }
                    
                    .publication-bar {
                    display: flex;
                    align-items: center;
                    justify-content: space-between;
                    gap: 30px;
                    padding: 18px 30px;
                    border-bottom: 4px solid #1e708b;
                    background: #f8fbfc;
                    font-family: Arial, Helvetica, sans-serif;
                    }
                    
                    .journal-brand {
                    flex: 0 0 auto;
                    color: #174f62;
                    font-size: 1.15rem;
                    font-weight: 700;
                    letter-spacing: 0.025em;
                    white-space: nowrap;
                    }
                    
                    .publication-metadata {
                    display: flex;
                    flex: 1 1 auto;
                    align-items: center;
                    justify-content: flex-end;
                    min-width: 0;
                    font-size: 0.82rem;
                    line-height: 1.35;
                    }
                    
                    .publication-item {
                    display: flex;
                    align-items: baseline;
                    gap: 5px;
                    padding: 3px 12px;
                    border-left: 1px solid #b9cbd2;
                    white-space: nowrap;
                    }
                    
                    .publication-item:first-child {
                    border-left: 0;
                    }
                    
                    .publication-label {
                    color: #4c5b61;
                    font-weight: 700;
                    }
                    
                    .publication-value {
                    color: #111111;
                    }
                    
                    .publication-doi {
                    white-space: normal;
                    overflow-wrap: anywhere;
                    }
                    
                    .article-grid {
                    display: grid;
                    grid-template-columns:
                    minmax(215px, 270px)
                    minmax(0, 1fr);
                    align-items: start;
                    gap: 42px;
                    padding: 38px 42px 52px;
                    }
                    
                    .citation-sidebar {
                    position: sticky;
                    top: 20px;
                    padding: 22px 20px;
                    border: 1px solid #cbd8dd;
                    border-top: 5px solid #1e708b;
                    background: #f6f9fa;
                    font-family: Arial, Helvetica, sans-serif;
                    font-size: 0.84rem;
                    line-height: 1.5;
                    overflow-wrap: anywhere;
                    }
                    
                    .citation-sidebar h2 {
                    margin: 0 0 15px;
                    color: #174f62;
                    font-size: 1.08rem;
                    line-height: 1.25;
                    }
                    
                    .citation-text {
                    margin: 0 0 22px;
                    text-align: left;
                    }
                    
                    .peer-review {
                    margin: 20px 0;
                    padding-top: 18px;
                    border-top: 1px solid #cbd8dd;
                    font-weight: 600;
                    }
                    
                    .editorial-dates {
                    margin-top: 18px;
                    padding-top: 15px;
                    border-top: 1px solid #cbd8dd;
                    }
                    
                    .editorial-date {
                    margin: 0 0 10px;
                    }
                    
                    .article-main {
                    min-width: 0;
                    }
                    
                    .article-header {
                    padding-bottom: 28px;
                    border-bottom: 1px solid #d7d7d7;
                    }
                    
                    .article-title {
                    margin: 0 0 16px;
                    color: #202020;
                    font-size: clamp(1.75rem, 3vw, 2.35rem);
                    line-height: 1.22;
                    }
                    
                    .translated-title {
                    margin: 0 0 28px;
                    color: #4c4c4c;
                    font-size: 1.18rem;
                    font-style: italic;
                    font-weight: normal;
                    line-height: 1.4;
                    }
                    
                    .authors {
                    margin: 25px 0 20px;
                    font-family: Arial, Helvetica, sans-serif;
                    }
                    
                    .author {
                    margin-bottom: 10px;
                    }
                    
                    .author-name {
                    font-weight: 700;
                    }
                    
                    .author-detail {
                    display: inline-block;
                    margin-left: 10px;
                    color: #3f555d;
                    font-size: 0.88rem;
                    }
                    
                    .affiliations {
                    margin: 20px 0 32px;
                    padding-left: 25px;
                    color: #383838;
                    font-size: 0.94rem;
                    }
                    
                    .affiliations li {
                    margin-bottom: 8px;
                    }
                    
                    .abstract {
                    margin: 34px 0;
                    padding: 22px 25px;
                    background: #f8f8f8;
                    border-left: 4px solid #6a909d;
                    }
                    
                    .abstract h2 {
                    margin: 0 0 12px;
                    font-size: 1.3rem;
                    }
                    
                    .abstract p {
                    margin-bottom: 0;
                    }
                    
                    .keywords {
                    margin: 16px 0;
                    font-family: Arial, Helvetica, sans-serif;
                    font-size: 0.94rem;
                    }
                    
                    main {
                    margin-top: 42px;
                    }
                    
                    section {
                    margin: 38px 0;
                    }
                    
                    section section {
                    margin: 28px 0 28px 24px;
                    }
                    
                    h2 {
                    color: #174f62;
                    font-size: 1.55rem;
                    line-height: 1.3;
                    }
                    
                    h3 {
                    color: #294c58;
                    font-size: 1.24rem;
                    line-height: 1.35;
                    }
                    
                    p {
                    margin: 0 0 18px;
                    text-align: justify;
                    }
                    
                    .tables-section {
                    margin-top: 54px;
                    padding-top: 18px;
                    border-top: 1px solid #d7d7d7;
                    }
                    
                    .tables-section-title {
                    margin-bottom: 28px;
                    }
                    
                    .table-wrapper {
                    width: 100%;
                    margin: 34px 0 48px;
                    overflow-x: auto;
                    border: 1px solid #c6d1d5;
                    background: #ffffff;
                    }
                    
                    .table-heading {
                    padding: 15px 18px;
                    border-bottom: 1px solid #c6d1d5;
                    background: #edf3f5;
                    }
                    
                    .table-label {
                    display: inline;
                    color: #174f62;
                    font-family: Arial, Helvetica, sans-serif;
                    font-weight: 700;
                    }
                    
                    .table-caption {
                    display: inline;
                    margin-left: 7px;
                    color: #283b42;
                    font-style: italic;
                    }
                    
                    table {
                    width: 100%;
                    min-width: 1150px;
                    border-collapse: collapse;
                    font-family: Arial, Helvetica, sans-serif;
                    font-size: 0.78rem;
                    line-height: 1.35;
                    }
                    
                    thead {
                    display: table-header-group;
                    }
                    
                    tbody {
                    display: table-row-group;
                    }
                    
                    th,
                    td {
                    padding: 9px 8px;
                    border: 1px solid #aebcc1;
                    text-align: left;
                    vertical-align: top;
                    overflow-wrap: anywhere;
                    }
                    
                    th {
                    color: #173f4c;
                    background: #dfeaec;
                    font-weight: 700;
                    }
                    
                    tbody tr:nth-child(even) {
                    background: #f7f9fa;
                    }
                    
                    tbody tr:hover {
                    background: #eef5f7;
                    }
                    
                    .figures {
                    margin-top: 50px;
                    padding-top: 18px;
                    border-top: 1px solid #d7d7d7;
                    }
                    
                    .article-figure {
                    margin: 36px auto;
                    padding: 20px;
                    border: 1px solid #d5d5d5;
                    background: #fafafa;
                    text-align: center;
                    }
                    
                    .article-figure img {
                    display: block;
                    width: auto;
                    max-width: 100%;
                    height: auto;
                    margin: 0 auto;
                    }
                    
                    .article-figure figcaption {
                    margin-top: 14px;
                    font-family: Arial, Helvetica, sans-serif;
                    font-size: 0.92rem;
                    line-height: 1.45;
                    text-align: left;
                    }
                    
                    .figure-label {
                    font-weight: 700;
                    }
                    
                    .figure-caption {
                    margin-left: 6px;
                    }
                    
                    .references {
                    margin-top: 52px;
                    padding-top: 20px;
                    border-top: 1px solid #d7d7d7;
                    }
                    
                    .references-list {
                    padding-left: 30px;
                    }
                    
                    .references-list li {
                    margin-bottom: 16px;
                    text-align: justify;
                    }
                    
                    .license-footer {
                    padding: 28px 42px 34px;
                    border-top: 4px solid #1e708b;
                    background: #f2f6f7;
                    font-family: Arial, Helvetica, sans-serif;
                    font-size: 0.9rem;
                    line-height: 1.55;
                    }
                    
                    .copyright-statement {
                    margin-bottom: 13px;
                    font-weight: 700;
                    }
                    
                    .license-statement {
                    margin-bottom: 0;
                    }
                    
                    @media (max-width: 1050px) {
                    .publication-bar {
                    align-items: flex-start;
                    flex-direction: column;
                    gap: 12px;
                    }
                    
                    .publication-metadata {
                    flex-wrap: wrap;
                    justify-content: flex-start;
                    }
                    
                    .publication-item {
                    border-left: 0;
                    border-right: 1px solid #b9cbd2;
                    }
                    }
                    
                    @media (max-width: 820px) {
                    .article-container {
                    width: 100%;
                    margin: 0;
                    box-shadow: none;
                    }
                    
                    .article-grid {
                    grid-template-columns: 1fr;
                    gap: 28px;
                    padding: 28px 24px 42px;
                    }
                    
                    .citation-sidebar {
                    position: static;
                    }
                    
                    .license-footer {
                    padding: 26px 24px;
                    }
                    
                    table {
                    min-width: 1050px;
                    }
                    }
                    
                    @media (max-width: 600px) {
                    body {
                    font-size: 16px;
                    }
                    
                    .publication-bar {
                    padding: 16px 20px;
                    }
                    
                    .publication-metadata {
                    display: block;
                    width: 100%;
                    }
                    
                    .publication-item {
                    display: block;
                    padding: 5px 0;
                    border-right: 0;
                    }
                    
                    .article-grid {
                    padding: 24px 18px 36px;
                    }
                    
                    .article-title {
                    font-size: 1.65rem;
                    }
                    
                    section section {
                    margin-left: 8px;
                    }
                    
                    table {
                    min-width: 950px;
                    font-size: 0.73rem;
                    }
                    }
                
                    /* FACPyA RED WHITE OVERRIDES */

                    :root {
                    --facpya-red: #9b1c20;
                    --facpya-red-dark: #6f1014;
                    --facpya-red-pale: #fff6f6;
                    --facpya-border: #e4bfc1;
                    --facpya-white: #ffffff;
                    }

                    body {
                    background: #f1f1f1;
                    color: #242424;
                    }

                    a {
                    color: var(--facpya-red-dark);
                    }

                    a:hover {
                    color: var(--facpya-red);
                    }

                    .article-container {
                    border-top: 7px solid var(--facpya-red-dark);
                    }

                    .publication-bar {
                    border-bottom: 4px solid var(--facpya-white);
                    background:
                    linear-gradient(
                    135deg,
                    var(--facpya-red-dark),
                    var(--facpya-red)
                    );
                    color: var(--facpya-white);
                    }

                    .journal-brand,
                    .publication-label,
                    .publication-value {
                    color: var(--facpya-white);
                    }

                    .publication-item {
                    border-left-color:
                    rgba(255, 255, 255, 0.40);
                    }

                    .citation-sidebar {
                    border-color: var(--facpya-border);
                    border-top-color: var(--facpya-red);
                    background: var(--facpya-white);
                    }

                    .citation-sidebar h2,
                    .article-title,
                    h2 {
                    color: var(--facpya-red-dark);
                    }

                    h3 {
                    color: var(--facpya-red);
                    }

                    .authors {
                    padding: 18px 20px;
                    border-left: 5px solid var(--facpya-red);
                    background: var(--facpya-red-pale);
                    }

                    .author-name {
                    color: var(--facpya-red-dark);
                    }

                    .abstract {
                    border-left-color: var(--facpya-red);
                    background: var(--facpya-red-pale);
                    }

                    .abstract h2,
                    .keywords strong {
                    color: var(--facpya-red-dark);
                    }

                    .keywords {
                    padding: 12px 15px;
                    border-left: 4px solid var(--facpya-red);
                    background: #fffafa;
                    }

                    .tables-section,
                    .figures,
                    .references {
                    border-top: 3px solid var(--facpya-red);
                    }

                    .table-wrapper {
                    border-color: var(--facpya-border);
                    }

                    .table-heading {
                    border-bottom-color:
                    var(--facpya-red-dark);
                    background:
                    linear-gradient(
                    135deg,
                    var(--facpya-red-dark),
                    var(--facpya-red)
                    );
                    }

                    .table-label,
                    .table-caption {
                    color: var(--facpya-white);
                    }

                    th {
                    color: var(--facpya-white);
                    background: var(--facpya-red-dark);
                    }

                    th,
                    td {
                    border-color: #d8adaf;
                    }

                    tbody tr:nth-child(even) {
                    background: var(--facpya-red-pale);
                    }

                    tbody tr:hover {
                    background: #f8e4e5;
                    }

                    .article-figure {
                    border-color: var(--facpya-border);
                    border-top: 5px solid var(--facpya-red);
                    background: var(--facpya-white);
                    }

                    .figure-label {
                    color: var(--facpya-red-dark);
                    }

                    .references-list li::marker {
                    color: var(--facpya-red);
                    font-weight: 700;
                    }

                    .license-footer {
                    border-top-color: var(--facpya-red-dark);
                    background:
                    linear-gradient(
                    135deg,
                    var(--facpya-red-dark),
                    var(--facpya-red)
                    );
                    color: var(--facpya-white);
                    }

                    .license-footer a {
                    color: var(--facpya-white);
                    }

                
                    /* CORRECCION SIN BARRA LATERAL */

                    .article-grid &gt; .article-main:only-child {
                    grid-column: 1 / -1;
                    width: 100%;
                    }

                    .article-grid &gt; .article-main:only-child
                    .article-title {
                    max-width: 1100px;
                    }

                    .article-grid &gt; .article-main:only-child
                    .translated-title {
                    max-width: 1100px;
                    }

                </style>
            </head>
            
            <body>
                <article class="article-container">
                    
                    <header class="publication-bar">
                        <div class="journal-brand">
                            <xsl:value-of select="article/front/journal-meta/journal-title-group/journal-title"/>
                        </div>
                        
                        <div class="publication-metadata">
                            <xsl:if test="$article-meta/article-id[@pub-id-type='doi']">
                                
                                <div class="publication-item">
                                    <span class="publication-label">
                                        DOI:
                                    </span>
                                    
                                    <span class="publication-value publication-doi">
                                        <xsl:value-of select="$article-meta/article-id[@pub-id-type='doi']"/>
                                    </span>
                                </div>
                            </xsl:if>
                            
                            <xsl:if test="$article-meta/volume">
                                <div class="publication-item">
                                    <span class="publication-label">
                                        Volumen:
                                    </span>
                                    
                                    <span class="publication-value">
                                        <xsl:value-of select="$article-meta/volume"/>
                                    </span>
                                </div>
                            </xsl:if>
                            
                            <xsl:if test="$article-meta/issue">
                                <div class="publication-item">
                                    <span class="publication-label">
                                        Número:
                                    </span>
                                    
                                    <span class="publication-value">
                                        <xsl:value-of select="$article-meta/issue"/>
                                    </span>
                                </div>
                            </xsl:if>
                            
                            <xsl:if test="$article-meta/fpage">
                                <div class="publication-item">
                                    <span class="publication-label">
                                        Páginas:
                                    </span>
                                    
                                    <span class="publication-value">
                                        <xsl:value-of select="$article-meta/fpage"/>
                                        
                                        <xsl:text>–</xsl:text>
                                        
                                        <xsl:value-of select="$article-meta/lpage"/>
                                    </span>
                                </div>
                            </xsl:if>
                            
                            <xsl:if test="$article-meta/pub-date[@date-type='pub']">
                                
                                <div class="publication-item">
                                    <span class="publication-label">
                                        Publicado:
                                    </span>
                                    
                                    <span class="publication-value">
                                        <xsl:value-of select="$article-meta/pub-date[@date-type='pub']/day"/>
                                        
                                        <xsl:text>/</xsl:text>
                                        
                                        <xsl:value-of select="$article-meta/pub-date[@date-type='pub']/month"/>
                                        
                                        <xsl:text>/</xsl:text>
                                        
                                        <xsl:value-of select="$article-meta/pub-date[@date-type='pub']/year"/>
                                    </span>
                                </div>
                            </xsl:if>
                        </div>
                    </header>
                    
                    <div class="article-grid">
                        
                        <xsl:if test="                                 $editorial-meta[                                     meta-name='citation-text'                                     or meta-name='peer-review-text'                                     or meta-name='received-text'                                     or meta-name='accepted-text'                                     or meta-name='published-text'                                 ]                             ">
                            
                            <aside class="citation-sidebar">
                                <h2>
                                    <xsl:choose>
                                        <xsl:when test="$editorial-meta[meta-name='citation-title']/meta-value">
                                            
                                            <xsl:value-of select="$editorial-meta[meta-name='citation-title']/meta-value"/>
                                        </xsl:when>
                                        
                                        <xsl:otherwise>
                                            <xsl:text>Cómo citar</xsl:text>
                                        </xsl:otherwise>
                                    </xsl:choose>
                                </h2>
                                
                                <xsl:if test="$editorial-meta[meta-name='citation-text']/meta-value">
                                    
                                    <p class="citation-text">
                                        <xsl:variable name="citation-render-text" select="string($editorial-meta[meta-name='citation-text']/meta-value)"/>

                                        <xsl:choose>

                                            <xsl:when test="contains($citation-render-text, 'Vinculat&#xE9;gica EFAN, ')">
                                                <xsl:value-of select="substring-before($citation-render-text, 'Vinculat&#xE9;gica EFAN, ')"/>
                                                <em>
                                                    <xsl:text>Vinculat&#xE9;gica EFAN</xsl:text>
                                                </em>
                                                <xsl:text>, </xsl:text>
                                                <xsl:variable name="after-journal" select="substring-after($citation-render-text, 'Vinculat&#xE9;gica EFAN, ')"/>
                                                <xsl:choose>
                                                    <xsl:when test="contains($after-journal, ',')">
                                                        <em>
                                                            <xsl:value-of select="normalize-space(substring-before($after-journal, ','))"/>
                                                        </em>
                                                        <xsl:text>,</xsl:text>
                                                        <xsl:value-of select="substring-after($after-journal, ',')"/>
                                                    </xsl:when>
                                                    <xsl:otherwise>
                                                        <xsl:value-of select="$after-journal"/>
                                                    </xsl:otherwise>
                                                </xsl:choose>
                                            </xsl:when>

                                            <xsl:when test="contains($citation-render-text, 'VinculaT&#xE9;gica EFAN, ')">
                                                <xsl:value-of select="substring-before($citation-render-text, 'VinculaT&#xE9;gica EFAN, ')"/>
                                                <em>
                                                    <xsl:text>VinculaT&#xE9;gica EFAN</xsl:text>
                                                </em>
                                                <xsl:text>, </xsl:text>
                                                <xsl:variable name="after-journal-old" select="substring-after($citation-render-text, 'VinculaT&#xE9;gica EFAN, ')"/>
                                                <xsl:choose>
                                                    <xsl:when test="contains($after-journal-old, ',')">
                                                        <em>
                                                            <xsl:value-of select="normalize-space(substring-before($after-journal-old, ','))"/>
                                                        </em>
                                                        <xsl:text>,</xsl:text>
                                                        <xsl:value-of select="substring-after($after-journal-old, ',')"/>
                                                    </xsl:when>
                                                    <xsl:otherwise>
                                                        <xsl:value-of select="$after-journal-old"/>
                                                    </xsl:otherwise>
                                                </xsl:choose>
                                            </xsl:when>

                                            <xsl:otherwise>
                                                <xsl:value-of select="$citation-render-text"/>
                                            </xsl:otherwise>
                                        </xsl:choose>
                                    </p>
                                </xsl:if>
                                
                                <xsl:if test="$editorial-meta[meta-name='peer-review-text']/meta-value">
                                    
                                    <p class="peer-review">
                                        <xsl:value-of select="$editorial-meta[meta-name='peer-review-text']/meta-value"/>
                                    </p>
                                </xsl:if>
                                
                                <xsl:if test="                                         $editorial-meta[meta-name='received-text']/meta-value                                         or $editorial-meta[meta-name='accepted-text']/meta-value                                         or $editorial-meta[meta-name='published-text']/meta-value                                     ">
                                    
                                    <div class="editorial-dates">
                                        <xsl:if test="$editorial-meta[meta-name='received-text']/meta-value">
                                            
                                            <p class="editorial-date">
                                                <xsl:value-of select="$editorial-meta[meta-name='received-text']/meta-value"/>
                                            </p>
                                        </xsl:if>
                                        
                                        <xsl:if test="$editorial-meta[meta-name='accepted-text']/meta-value">
                                            
                                            <p class="editorial-date">
                                                <xsl:value-of select="$editorial-meta[meta-name='accepted-text']/meta-value"/>
                                            </p>
                                        </xsl:if>
                                        
                                        <xsl:if test="$editorial-meta[meta-name='published-text']/meta-value">
                                            
                                            <p class="editorial-date">
                                                <xsl:value-of select="$editorial-meta[meta-name='published-text']/meta-value"/>
                                            </p>
                                        </xsl:if>
                                    </div>
                                </xsl:if>
                            </aside>
                        </xsl:if>
                        
                        <div class="article-main">
                            <header class="article-header">
                                <h1 class="article-title">
                                    <xsl:value-of select="$article-meta/title-group/article-title"/>
                                </h1>
                                
                                <xsl:if test="$article-meta/title-group/trans-title-group/trans-title">
                                    
                                    <h2 class="translated-title">
                                        <xsl:value-of select="$article-meta/title-group/trans-title-group/trans-title"/>
                                    </h2>
                                </xsl:if>
                                
                                <div class="authors">
                                    <xsl:apply-templates select="$article-meta/contrib-group/contrib" mode="author"/>
                                </div>
                                
                                <xsl:if test="$article-meta/aff">
                                    <ol class="affiliations">
                                        <xsl:apply-templates select="$article-meta/aff" mode="affiliation"/>
                                    </ol>
                                </xsl:if>
                                
                                <xsl:apply-templates select="$article-meta/abstract"/>
                                
                                <xsl:apply-templates select="$article-meta/trans-abstract"/>
                                
                                <xsl:apply-templates select="$article-meta/kwd-group"/>
                            </header>
                            
                            <main>
                                <xsl:apply-templates select="article/body/sec"/>
                            </main>
                            
                            <xsl:apply-templates select="article/floats-group"/>
                            
                            <xsl:apply-templates select="article/back/ref-list"/>
                        </div>
                    </div>
                    
                    <xsl:apply-templates select="$article-meta/permissions"/>
                </article>
            </body>
        </html>
    </xsl:template>
    
    <xsl:template match="contrib" mode="author">
        
        <div class="author">
            <span class="author-name">
                <xsl:value-of select="name/given-names"/>
                
                <xsl:text> </xsl:text>
                
                <xsl:value-of select="name/surname"/>
            </span>
            
            <xsl:for-each select="xref[@ref-type='aff']">
                <sup>
                    <xsl:value-of select="."/>
                </sup>
            </xsl:for-each>
            
            <xsl:if test="contrib-id[@contrib-id-type='orcid']">
                <span class="author-detail">
                    <xsl:text>ORCID: </xsl:text>
                    
                    <xsl:value-of select="contrib-id[@contrib-id-type='orcid']"/>
                </span>
            </xsl:if>
            
            <xsl:if test="email">
                <span class="author-detail">
                    <xsl:value-of select="email"/>
                </span>
            </xsl:if>
        </div>
    </xsl:template>
    
    <xsl:template match="aff" mode="affiliation">
        
        <li>
            <xsl:value-of select="institution[@content-type='original']"/>
            
            <xsl:if test="country">
                <xsl:text>. </xsl:text>
                
                <xsl:value-of select="country"/>
            </xsl:if>
        </li>
    </xsl:template>
    
    <xsl:template match="abstract">
        <section class="abstract">
            <h2>
                <xsl:value-of select="title"/>
            </h2>
            
            <xsl:apply-templates select="p"/>
        </section>
    </xsl:template>
    
    <xsl:template match="trans-abstract">
        <section class="abstract">
            <h2>
                <xsl:value-of select="title"/>
            </h2>
            
            <xsl:apply-templates select="p"/>
        </section>
    </xsl:template>
    
    <xsl:template match="kwd-group">
        <div class="keywords">
            <strong>
                <xsl:value-of select="title"/>
                
                <xsl:text>: </xsl:text>
            </strong>
            
            <xsl:for-each select="kwd">
                <xsl:value-of select="."/>
                
                <xsl:if test="position() != last()">
                    <xsl:text>, </xsl:text>
                </xsl:if>
            </xsl:for-each>
        </div>
    </xsl:template>
    
    <xsl:template match="sec">
        <section id="{@id}">
            <xsl:choose>
                <xsl:when test="parent::body">
                    <h2>
                        <xsl:value-of select="title"/>
                    </h2>
                </xsl:when>
                
                <xsl:otherwise>
                    <h3>
                        <xsl:value-of select="title"/>
                    </h3>
                </xsl:otherwise>
            </xsl:choose>
            
            <xsl:apply-templates select="p | table-wrap | fig | sec | list | disp-formula"/>
        </section>
    </xsl:template>
    
    <xsl:template match="p">
        <p>
            <xsl:apply-templates/>
        </p>
    </xsl:template>
    
    <xsl:template match="floats-group">
        <xsl:if test="table-wrap">
            <section class="tables-section">
                <h2 class="tables-section-title">
                    Tablas
                </h2>
                
                <xsl:apply-templates select="table-wrap"/>
            </section>
        </xsl:if>
        
        <xsl:if test="fig">
            <section class="figures">
                <h2>Figuras</h2>
                
                <xsl:apply-templates select="fig"/>
            </section>
        </xsl:if>
    </xsl:template>
    
    <xsl:template match="table-wrap">
        <div class="table-wrapper" id="{@id}">
            <div class="table-heading">
                <xsl:if test="label">
                    <span class="table-label">
                        <xsl:value-of select="label"/>
                    </span>
                </xsl:if>
                
                <xsl:if test="caption/title">
                    <span class="table-caption">
                        <xsl:value-of select="caption/title"/>
                    </span>
                </xsl:if>
            </div>
            
            <xsl:apply-templates select="table"/>
        </div>
    </xsl:template>
    
    <xsl:template match="table">
        <table>
            <xsl:apply-templates/>
        </table>
    </xsl:template>
    
    <xsl:template match="thead">
        <thead>
            <xsl:apply-templates/>
        </thead>
    </xsl:template>
    
    <xsl:template match="tbody">
        <tbody>
            <xsl:apply-templates/>
        </tbody>
    </xsl:template>
    
    <xsl:template match="tr">
        <tr>
            <xsl:apply-templates/>
        </tr>
    </xsl:template>
    
    <xsl:template match="th">
        <th>
            <xsl:apply-templates/>
        </th>
    </xsl:template>
    
    <xsl:template match="td">
        <td>
            <xsl:apply-templates/>
        </td>
    </xsl:template>
    
    <xsl:template match="fig">
        <figure class="article-figure" id="{@id}">
            <xsl:if test="graphic">
                <img src="{graphic/@xlink:href}">
                    <xsl:attribute name="alt">
                        <xsl:choose>
                            <xsl:when test="normalize-space(alt-text)">
                                <xsl:value-of select="alt-text"/>
                            </xsl:when>
                            
                            <xsl:when test="normalize-space(caption/title)">
                                <xsl:value-of select="caption/title"/>
                            </xsl:when>
                            
                            <xsl:when test="normalize-space(label)">
                                <xsl:value-of select="label"/>
                            </xsl:when>
                            
                            <xsl:otherwise>
                                <xsl:text>Figura del artículo</xsl:text>
                            </xsl:otherwise>
                        </xsl:choose>
                    </xsl:attribute>
                </img>
            </xsl:if>
            
            <xsl:if test="label or caption/title">
                <figcaption>
                    <xsl:if test="label">
                        <span class="figure-label">
                            <xsl:value-of select="label"/>
                        </span>
                    </xsl:if>
                    
                    <xsl:if test="caption/title">
                        <span class="figure-caption">
                            <xsl:value-of select="caption/title"/>
                        </span>
                    </xsl:if>
                </figcaption>
            </xsl:if>
        </figure>
    </xsl:template>
    
    <xsl:template match="ref-list">
        <section class="references">
            <h2>
                <xsl:value-of select="title"/>
            </h2>
            
            <ol class="references-list">
                <xsl:apply-templates select="ref"/>
            </ol>
        </section>
    </xsl:template>
    
    <xsl:template match="ref">
        <li id="{@id}">
            <xsl:value-of select="mixed-citation"/>
        </li>
    </xsl:template>
    
    <xsl:template match="permissions">
        <footer class="license-footer">
            <xsl:if test="copyright-statement">
                <div class="copyright-statement">
                    <xsl:value-of select="copyright-statement"/>
                </div>
            </xsl:if>
            
            <xsl:if test="license">
                <p class="license-statement">
                    <xsl:value-of select="license/license-p"/>
                    
                    <xsl:if test="license/@xlink:href">
                        <xsl:text> </xsl:text>
                        
                        <a href="{license/@xlink:href}">
                            <xsl:value-of select="license/@xlink:href"/>
                        </a>
                    </xsl:if>
                </p>
            </xsl:if>
        </footer>
    </xsl:template>
    
<xsl:template match="list"><xsl:choose><xsl:when test="@list-type='order'"><ol class="article-list article-list-ordered"><xsl:apply-templates select="list-item"/></ol></xsl:when><xsl:otherwise><ul class="article-list article-list-bullet"><xsl:apply-templates select="list-item"/></ul></xsl:otherwise></xsl:choose></xsl:template><xsl:template match="list-item"><li><xsl:apply-templates/></li></xsl:template><xsl:template match="disp-formula"><div class="display-formula" style="display:grid;grid-template-columns:minmax(0,1fr) auto;align-items:center;gap:16px;margin:1.25em 0;overflow-x:auto;"><div class="formula-math" style="text-align:center;min-width:0;"><xsl:apply-templates select="*[local-name()='math' and namespace-uri()='http://www.w3.org/1998/Math/MathML']" mode="html-math"/></div><xsl:if test="label"><span class="formula-label" style="white-space:nowrap;"><xsl:value-of select="label"/></span></xsl:if></div></xsl:template><xsl:template match="*" mode="html-math"><xsl:element name="{local-name()}"><xsl:copy-of select="@*"/><xsl:apply-templates select="node()" mode="html-math"/></xsl:element></xsl:template></xsl:stylesheet>