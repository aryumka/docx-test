package aryumka.docxexample;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import java.io.*;
import java.math.BigInteger;
import java.util.*;

public class WordMerger {

    public static InputStream merge(final List<InputStream> streams) throws Docx4JException, IOException {
        WordprocessingMLPackage target = null;
        final File generated = File.createTempFile("generated", ".docx");

        List<String> existingStyleIds = new ArrayList<>();
        Map<String, String> styleIdMap = new HashMap<>();
        Map<BigInteger, BigInteger> numIdMap = new HashMap<>();
        BigInteger maxNumId = BigInteger.ZERO;

        for (InputStream is : streams) {
            if (is != null) {
                if (target == null) {
                    // 첫 번째 문서를 로드합니다.
                    target = WordprocessingMLPackage.load(is);

                    // 기존 스타일 ID와 최대 numId를 가져옵니다.
                    existingStyleIds.addAll(getStyleIds(target));
                    maxNumId = getMaxNumId(target);
                } else {
                    WordprocessingMLPackage source = WordprocessingMLPackage.load(is);

                    // 페이지 나누기 추가
                    addPageBreak(target);

                    // 스타일 병합
                    renameAndMergeStyles(target, source, existingStyleIds, styleIdMap);

                    // 번호 매기기 병합
                    maxNumId = mergeNumbering(target, source, numIdMap, maxNumId);

                    // 관계 병합
                    mergeRelationships(target, source);

                    // 콘텐츠 병합
                    mergeContent(target, source, styleIdMap, numIdMap);
                }
            }
        }

        if (target != null) {
            target.save(generated);
            return new FileInputStream(generated);
        } else {
            return null;
        }
    }

    // 페이지 나누기 추가 메서드
    private static void addPageBreak(WordprocessingMLPackage wordPackage) {
        Br breakObj = new Br();
        breakObj.setType(STBrType.PAGE);

        R run = Context.getWmlObjectFactory().createR();
        run.getContent().add(breakObj);

        P para = Context.getWmlObjectFactory().createP();
        para.getContent().add(run);

        wordPackage.getMainDocumentPart().getContent().add(para);
    }

    private static List<String> getStyleIds(WordprocessingMLPackage pack) {
        List<String> styleIds = new ArrayList<>();
        StyleDefinitionsPart stylesPart = pack.getMainDocumentPart().getStyleDefinitionsPart();
        if (stylesPart != null) {
            for (Style style : stylesPart.getJaxbElement().getStyle()) {
                styleIds.add(style.getStyleId());
            }
        }
        return styleIds;
    }

    private static BigInteger getMaxNumId(WordprocessingMLPackage pack) {
        BigInteger maxNumId = BigInteger.ZERO;
        NumberingDefinitionsPart numberingPart = pack.getMainDocumentPart().getNumberingDefinitionsPart();
        if (numberingPart != null) {
            for (Numbering.Num num : numberingPart.getJaxbElement().getNum()) {
                if (num.getNumId().compareTo(maxNumId) > 0) {
                    maxNumId = num.getNumId();
                }
            }
        }
        return maxNumId;
    }

    private static void renameAndMergeStyles(WordprocessingMLPackage target, WordprocessingMLPackage source,
                                             List<String> existingStyleIds, Map<String, String> styleIdMap) throws Docx4JException {
        StyleDefinitionsPart sourceStylesPart = source.getMainDocumentPart().getStyleDefinitionsPart();
        StyleDefinitionsPart targetStylesPart = target.getMainDocumentPart().getStyleDefinitionsPart();

        if (sourceStylesPart != null && targetStylesPart != null) {
            for (Style style : sourceStylesPart.getJaxbElement().getStyle()) {
                String originalId = style.getStyleId();
                String newId = originalId;
                while (existingStyleIds.contains(newId)) {
                    newId = newId + "_copy";
                }
                if (!originalId.equals(newId)) {
                    styleIdMap.put(originalId, newId);
                    style.setStyleId(newId);
                    if (style.getBasedOn() != null) {
                        style.getBasedOn().setVal(renameStyleId(style.getBasedOn().getVal(), existingStyleIds, styleIdMap));
                    }
                    if (style.getNext() != null) {
                        style.getNext().setVal(renameStyleId(style.getNext().getVal(), existingStyleIds, styleIdMap));
                    }
                    if (style.getLink() != null) {
                        style.getLink().setVal(renameStyleId(style.getLink().getVal(), existingStyleIds, styleIdMap));
                    }
                }
                existingStyleIds.add(newId);
                targetStylesPart.getJaxbElement().getStyle().add(style);
            }
        }
    }

    private static String renameStyleId(String styleId, List<String> existingStyleIds, Map<String, String> styleIdMap) {
        if (styleIdMap.containsKey(styleId)) {
            return styleIdMap.get(styleId);
        }
        String newId = styleId;
        while (existingStyleIds.contains(newId)) {
            newId = newId + "_copy";
        }
        styleIdMap.put(styleId, newId);
        existingStyleIds.add(newId);
        return newId;
    }

    private static BigInteger mergeNumbering(WordprocessingMLPackage target, WordprocessingMLPackage source,
                                             Map<BigInteger, BigInteger> numIdMap, BigInteger maxNumId) throws Docx4JException {
        NumberingDefinitionsPart sourceNumberingPart = source.getMainDocumentPart().getNumberingDefinitionsPart();
        NumberingDefinitionsPart targetNumberingPart = target.getMainDocumentPart().getNumberingDefinitionsPart();

        if (sourceNumberingPart != null) {
            if (targetNumberingPart == null) {
                targetNumberingPart = new NumberingDefinitionsPart();
                target.getMainDocumentPart().addTargetPart(targetNumberingPart);
                targetNumberingPart.setJaxbElement(Context.getWmlObjectFactory().createNumbering());
            }

            // AbstractNum 병합
            for (Numbering.AbstractNum abstractNum : sourceNumberingPart.getJaxbElement().getAbstractNum()) {
                BigInteger newAbstractNumId = maxNumId.add(BigInteger.ONE);
                maxNumId = newAbstractNumId;
                abstractNum.setAbstractNumId(newAbstractNumId);
                targetNumberingPart.getJaxbElement().getAbstractNum().add(abstractNum);
            }

            // Num 병합
            for (Numbering.Num num : sourceNumberingPart.getJaxbElement().getNum()) {
                BigInteger oldNumId = num.getNumId();
                BigInteger newNumId = maxNumId.add(BigInteger.ONE);
                maxNumId = newNumId;
                numIdMap.put(oldNumId, newNumId);
                num.setNumId(newNumId);

                // abstractNumId 업데이트
                BigInteger oldAbstractNumId = num.getAbstractNumId().getVal();
                BigInteger newAbstractNumId = num.getAbstractNumId().getVal().add(maxNumId);
                num.getAbstractNumId().setVal(newAbstractNumId);

                targetNumberingPart.getJaxbElement().getNum().add(num);
            }
        }
        return maxNumId;
    }

    private static void mergeRelationships(WordprocessingMLPackage target, WordprocessingMLPackage source) throws Docx4JException {
        for (Relationship rel : source.getMainDocumentPart().getRelationshipsPart().getRelationships().getRelationship()) {
            if (!rel.getType().equals(Namespaces.STYLES)
                    && !rel.getType().equals(Namespaces.NUMBERING)) {
                Part part = source.getMainDocumentPart().getRelationshipsPart().getPart(rel);
                target.getMainDocumentPart().addTargetPart(part);
            }
        }
    }

    private static void mergeContent(WordprocessingMLPackage target, WordprocessingMLPackage source,
                                     Map<String, String> styleIdMap, Map<BigInteger, BigInteger> numIdMap) {
        List<Object> sourceContent = source.getMainDocumentPart().getContent();
        updateStyleAndNumIds(sourceContent, styleIdMap, numIdMap);
        target.getMainDocumentPart().getContent().addAll(sourceContent);
    }

    private static void updateStyleAndNumIds(List<Object> content,
                                             Map<String, String> styleIdMap, Map<BigInteger, BigInteger> numIdMap) {
        for (Object obj : content) {
            if (obj instanceof P) {
                P p = (P) obj;
                // 단락 스타일 업데이트
                if (p.getPPr() != null) {
                    if (p.getPPr().getPStyle() != null) {
                        String oldStyleId = p.getPPr().getPStyle().getVal();
                        if (styleIdMap.containsKey(oldStyleId)) {
                            p.getPPr().getPStyle().setVal(styleIdMap.get(oldStyleId));
                        }
                    }
                    // 번호 매기기 ID 업데이트
                    if (p.getPPr().getNumPr() != null && p.getPPr().getNumPr().getNumId() != null) {
                        BigInteger oldNumId = p.getPPr().getNumPr().getNumId().getVal();
                        if (numIdMap.containsKey(oldNumId)) {
                            p.getPPr().getNumPr().getNumId().setVal(numIdMap.get(oldNumId));
                        }
                    }
                }
                // 런 스타일 업데이트
                updateRunStyles(p.getContent(), styleIdMap);
            } else if (obj instanceof Tbl) {
                // 표 내의 스타일 업데이트
                Tbl tbl = (Tbl) obj;
                updateTableStyles(tbl, styleIdMap, numIdMap);
            }
        }
    }

    private static void updateRunStyles(List<Object> content, Map<String, String> styleIdMap) {
        for (Object obj : content) {
            if (obj instanceof R) {
                R r = (R) obj;
                if (r.getRPr() != null && r.getRPr().getRStyle() != null) {
                    String oldStyleId = r.getRPr().getRStyle().getVal();
                    if (styleIdMap.containsKey(oldStyleId)) {
                        r.getRPr().getRStyle().setVal(styleIdMap.get(oldStyleId));
                    }
                }
            } else if (obj instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) obj;
                if (element.getValue() instanceof CTSimpleField) {
                    CTSimpleField field = (CTSimpleField) element.getValue();
                    updateRunStyles(field.getContent(), styleIdMap);
                } else if (element.getValue() instanceof P) {
                    P p = (P) element.getValue();
                    updateRunStyles(p.getContent(), styleIdMap);
                }
            }
        }
    }

    private static void updateTableStyles(Tbl tbl, Map<String, String> styleIdMap, Map<BigInteger, BigInteger> numIdMap) {
        if (tbl.getTblPr() != null && tbl.getTblPr().getTblStyle() != null) {
            String oldStyleId = tbl.getTblPr().getTblStyle().getVal();
            if (styleIdMap.containsKey(oldStyleId)) {
                tbl.getTblPr().getTblStyle().setVal(styleIdMap.get(oldStyleId));
            }
        }
        for (Object rowObj : tbl.getContent()) {
            if (rowObj instanceof Tr) {
                Tr tr = (Tr) rowObj;
                for (Object cellObj : tr.getContent()) {
                    if (cellObj instanceof Tc) {
                        Tc tc = (Tc) cellObj;
                        updateStyleAndNumIds(tc.getContent(), styleIdMap, numIdMap);
                    }
                }
            }
        }
    }
}