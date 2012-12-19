all:	clean zip install

clean:
	unopkg remove loairviro.oxt
	rm loairviro.oxt
       
zip:
	zip -r loairviro.oxt \
		description.xml \
		META-INF/manifest.xml \
		registry/Addons.xcu \
		src/loairviro.py \
		package/desc_en.txt \
		package/license_en.txt

install:
	unopkg add loairviro.oxt