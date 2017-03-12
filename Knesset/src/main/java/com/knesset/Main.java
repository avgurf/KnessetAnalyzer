package com.knesset;
import com.google.gson.Gson;
import com.kennycason.kumo.CollisionMode;
import com.kennycason.kumo.WordCloud;
import com.kennycason.kumo.WordFrequency;
import com.kennycason.kumo.font.FontWeight;
import com.kennycason.kumo.font.KumoFont;
import com.kennycason.kumo.font.scale.LinearFontScalar;
import com.kennycason.kumo.nlp.FrequencyAnalyzer;
import com.mongodb.MongoClient;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.bson.Document;

import java.awt.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collection;

import static com.mongodb.client.model.Filters.eq;

/**
 * Created by Avi on 04/03/2017.
 */
public class Main {
	public static  final String STOP_WORD_FILE = "heb_stopwords.txt";

	public static Collection<String> loadStopWords() throws IOException {
		File file = new File(STOP_WORD_FILE);
		BufferedReader br = new BufferedReader(new FileReader(file));
		ArrayList<String> retVal = new ArrayList<String>();
		String word = "";
		while((word = br.readLine()) != null)
			retVal.add(word);

		return retVal;

	}
	public static void main(String[] args) {
		try
		{
			String in_dir = "in";
			String out_dir = "out";

			if(args.length >= 2) {
				in_dir = args[0];
				out_dir = args[1];
			}
			MongoClient mongoClient = new MongoClient( "localhost" , 27017);
			MongoDatabase database = mongoClient.getDatabase("knessetTranscripts");

			MongoCollection<Document> collection = database.getCollection("scripts");

			File folder = new File(in_dir);
			File[] listOfFiles = folder.listFiles();


			for (int i = 0; i < listOfFiles.length; i++) {
				Document doc = null;
				// check if file exists
				try {
					doc = collection.find(eq("fileName", listOfFiles[i].getName())).first();
					if(doc != null)
					{
						System.out.println("Found" + listOfFiles[i].getName() + " in DB:");
						System.out.println("\t" +doc.toJson());
						continue;
					}
				}
				catch (Exception e){
					e.printStackTrace();
				}
				if(doc == null)
					doc = new Document("fileName",listOfFiles[i].getName());

				if (listOfFiles[i].isFile()) {
					System.out.println("Analyzing File " + listOfFiles[i].getName());
					Document _doc = analyzeFile(listOfFiles[i].getName(), out_dir, in_dir ,doc);
					if(_doc != null)
					{
						System.out.println("\tAdding to MongoDB:" + _doc.toJson());
						collection.insertOne(_doc);
						//eq("fileName", listOfFiles[i].getName()),new Document("$set",_doc));
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static Document analyzeFile(String name, String out_path, String in_path, Document doc)
	{
		try {
			File file = new File(in_path + "\\"+ name);

			NPOIFSFileSystem fs = new NPOIFSFileSystem(file);
			WordExtractor extractor = new WordExtractor(fs.getRoot());
			String outfile = name.replace(".doc","") + ".txt";
			PrintWriter out = new PrintWriter(out_path + "\\" + outfile, "UTF-8");
			StringBuilder stringBuilder = new StringBuilder();
			for(String rawText : extractor.getMainTextboxText()) {
				String text = extractor.stripFields(rawText);
				// clear commas
				if(text.endsWith("-"+"\r\n"))
					text = text.replace("-\r\n","");
				System.out.print(text);
				out.write(new String(text.getBytes("UTF-8"), "UTF-8"));
				stringBuilder.append(text);
			}

			// add text to file field
			doc.append("content", stringBuilder.toString());

			final FrequencyAnalyzer frequencyAnalyzer = new FrequencyAnalyzer();
			frequencyAnalyzer.setWordFrequenciesToReturn(300);
			frequencyAnalyzer.setMinWordLength(4);
			frequencyAnalyzer.setStopWords(loadStopWords());
			frequencyAnalyzer.setCharacterEncoding("UTF-8");
			final java.util.List<WordFrequency> wordFrequencies = frequencyAnalyzer.load(out_path + "\\" + outfile);

			// Append the word frequency list filed
			doc.append("wordFrequency", new Gson().toJson(wordFrequencies));

			final Dimension dimension = new Dimension(500, 500);
			final WordCloud wordCloud = new WordCloud(dimension, CollisionMode.PIXEL_PERFECT);
			wordCloud.setPadding(2);
			wordCloud.setKumoFont(new KumoFont("Ariel", FontWeight.BOLD));
//			wordCloud.setBackground(new PixelBoundryBackground(getInputStream("backgrounds/whale_small.png")));
//			wordCloud.setColorPalette(new ColorPalette(new Color(0x4055F1), new Color(0x408DF1), new Color(0x40AAF1), new Color(0x40C5F1), new Color(0x40D3F1), new Color(0xFFFFFF)));
			wordCloud.setFontScalar(new LinearFontScalar(10, 72));
			wordCloud.build(wordFrequencies);
			String imageFilePath = out_path + "\\" + "wordcloud_small_"+ outfile.replace(".txt","")  +".png";
			wordCloud.writeToFile(imageFilePath);

			// Read the .png file and save it in the DB
			Path image = Paths.get(imageFilePath);
			doc.append("wordCloud", Files.readAllBytes(image));
			return  doc;
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
		return null;
	}
}
