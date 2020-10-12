import org.jsoup.Jsoup;

import java.io.IOException;
import java.io.PrintWriter;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;


public class QuestDataScraper {
    public static void main(String[] args) {
        Map<Integer, String> quests = new HashMap<>();
        Map<Integer, Map<String, Integer>> questItems = new HashMap<>();
        try {
            for (int i=50; i<=59; i++) {
                String webPage = String.format("https://shadowlands.wowhead.com/quests/min-level:%d/max-level:%d?filter=35;9;0", i, i);
                String html = Jsoup.connect(webPage).get().html();
                String pattern = "\\{\"category\":-?\\d+,(\"category2\":-?\\d+,)?(\"currencyrewards\":\\[\\[.*?\\]\\],)?(\"daily\":\\d+,)?\"id\":(\\d+).*?\"name\":\"(.*?)\"";
                collectQuestData(quests, html, pattern);
            }
            for (Map.Entry<Integer, String> entry : quests.entrySet()) {
                String webPage = String.format("https://shadowlands.wowhead.com/quest=%d", entry.getKey());
                String html = Jsoup.connect(webPage).get().html();
                String pattern = "(item=(\\d+))|(g_items.createIcon\\((\\d+))";
                collectQuestItemData(questItems, html, pattern, webPage);
            }
            writeQuestInfoToFile(quests, questItems);
        } catch (IOException e) {
            System.out.println(e.getStackTrace());
        }
    }

    private static void collectQuestData(Map<Integer, String> quests, String html, String pattern) {
        Matcher m = Pattern.compile(pattern).matcher(html);
        while (m.find()) {
            if (isLevelingQuest(m)) {
                quests.put(Integer.parseInt(m.group(4)), m.group(5));
            }
        }
    }

    private static boolean isLevelingQuest(Matcher m) {
        String questName = m.group(5).toLowerCase();
        if (questName.contains("deprecated") || questName.contains("unused") || questName.contains("dnt")
        || questName.contains("old not used") || questName.contains("shadowlands (51-59) e") || questName.contains("reuse me")
        || questName.contains("nyi") || questName.contains("professions - reuse") || questName.contains("tbd")) {
            return false;
        }
        return true;
    }

    private static void collectQuestItemData(Map<Integer, Map<String, Integer>> questItems, String html, String pattern, String oldWebPage) throws IOException {
        Matcher m = Pattern.compile(pattern).matcher(html);
        Set<Integer> itemIds = new HashSet<>();
        while (m.find()) {
            for (int i=2; i<=4; i+=2) {
                if (m.group(i) != null) {
                    itemIds.add(Integer.parseInt(m.group(i)));
                }
            }
        }
        for (Integer itemId : itemIds) {
            String webPage = String.format("https://shadowlands.wowhead.com/item=%d", itemId);
            String itemHtml = Jsoup.connect(webPage).get().html();
            if (itemHtml.toLowerCase().contains("cost") || itemHtml.toLowerCase().contains("plagueborn slime")
                    || itemHtml.toLowerCase().contains("argent dawn valor token")
                    || itemHtml.toLowerCase().contains("chromie's scroll")
                    || itemHtml.toLowerCase().contains("brimming stoneborn heart")
                    || itemHtml.toLowerCase().contains("memory of a vital sacrifice")) {
                continue;
            }
            Matcher m2 = Pattern.compile("(<h1(.|\\s)+?<\\/noscript>)").matcher(itemHtml);
            if (m2.find()) {
                String itemInfo = m2.group(1);
                if (itemInfo.toLowerCase().contains("toy") || itemInfo.toLowerCase().contains("unique-equipped")
                        || itemInfo.toLowerCase().contains("mount") || itemInfo.toLowerCase().contains("tabard")
                        || itemInfo.toLowerCase().contains("sell price")) {
                    continue;
                }
                Matcher m3 = Pattern.compile("spell=(\\d+)").matcher(itemInfo);
                if (!m3.find()) {
                    continue;
                }
                Map<String, Integer> questItemData = new HashMap<>();
                questItemData.put("spellID", Integer.parseInt(m3.group(1)));
                Matcher m4 = Pattern.compile("(\\d+)(?= Sec Cooldown)").matcher(itemInfo);
                if (m4.find()) {
                    questItemData.put("cooldown", Integer.parseInt(m4.group(1)));
                }
                questItems.put(itemId, questItemData);
                System.out.println(itemInfo.replaceAll("\\s", ""));
            }
        }
    }


    private static void writeQuestInfoToFile(Map<Integer, String> quests, Map<Integer, Map<String, Integer>> questItems) throws IOException {
        PrintWriter writer = new PrintWriter("questIDToName.txt", "UTF-8");
        writeQuestIDToName(quests, writer);
        writer.println();
        writeQuestNames(quests, writer);
        writer.println();
        writeQuestItems(questItems, writer);
        writer.close();
    }

    private static void writeQuestIDToName(Map<Integer, String> quests, PrintWriter writer) throws IOException {
        StringBuilder writeString = new StringBuilder("addonTable.questIDToName = {\r\n");
        for (Map.Entry<Integer, String> entry : quests.entrySet()) {
            writeString.append("    [" ).append(entry.getKey()).append("] = \"").append(entry.getValue().toLowerCase()).append("\",\r\n");
        }
        writeString.append("}");
        writer.println(writeString);
    }

    private static void writeQuestNames(Map<Integer, String> quests, PrintWriter writer) throws IOException {
        Set<String> questNames = new HashSet<>(quests.values());
        questNames = questNames.stream().map(String::toLowerCase).collect(Collectors.toSet());
        StringBuilder writeString = new StringBuilder("addonTable.questNames = {\r\n");
        questNames.stream().forEach(name->writeString.append("    [\"" ).append(name).append("\"] = true,\r\n"));
        writeString.append("}");
        writer.println(writeString);
    }

    private static void writeQuestItems(Map<Integer, Map<String, Integer>> questItems, PrintWriter writer) throws IOException {
        StringBuilder writeString = new StringBuilder("addonTable.questItems = {\r\n");
        for (Map.Entry<Integer, Map<String, Integer>> entry : questItems.entrySet()) {
            writeString.append("    [" ).append(entry.getKey()).append("] = {\r\n");
            writeString.append("        [\"count\"] = 0,\r\n");
            writeString.append("        [\"spellID\"] = ").append(entry.getValue().get("spellID")).append(",\r\n");
            Integer cooldown = entry.getValue().get("cooldown");
            writeString.append("        [\"cooldown\"] = ").append(cooldown != null ? cooldown : "nil").append("\r\n");
            writeString.append("    },\r\n");
        }
        writeString.append("}");
        writer.println(writeString);
    }
}
