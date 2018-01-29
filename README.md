# Web MarkDown 
* MarkDown Description Sample

***
## Link
```
[inline-style link](https://www.google.com)**

[inline-style link with title](https://www.google.com "Google's Homepage")

[reference-style link][Arbitrary case-insensitive reference text]

[relative reference to a repository file](../blob/master/LICENSE)

[You can use numbers for reference-style link definitions][1]
```
**[inline-style link](https://www.google.com)**

**[inline-style link with title](https://www.google.com "Google's Homepage")**

**[reference-style link][Arbitrary case-insensitive reference text]**

**[relative reference to a repository file](../blob/master/LICENSE)**

**[You can use numbers for reference-style link definitions][1]**

***
## List
```
1. First ordered list item
2. Another item
⋅⋅* Unordered sub-list. 
1. Actual numbers don't matter, just that it's a number
⋅⋅1. Ordered sub-list
4. And another item.

⋅⋅⋅You can have properly indented paragraphs within list items. Notice the blank line above, and the leading spaces (at least one, but we'll use three here to also align the raw Markdown).

⋅⋅⋅To have a line break without a paragraph, you will need to use two trailing spaces.⋅⋅
⋅⋅⋅Note that this line is separate, but within the same paragraph.⋅⋅
⋅⋅⋅(This is contrary to the typical GFM line break behaviour, where trailing spaces are not required.)

* Unordered list can use asterisks
- Or minuses
+ Or pluses
```
1. First ordered list item
2. Another item
⋅⋅* Unordered sub-list. 
1. Actual numbers don't matter, just that it's a number
⋅⋅1. Ordered sub-list
4. And another item.

⋅⋅⋅You can have properly indented paragraphs within list items. Notice the blank line above, and the leading spaces (at least one, but we'll use three here to also align the raw Markdown).

⋅⋅⋅To have a line break without a paragraph, you will need to use two trailing spaces.⋅⋅
⋅⋅⋅Note that this line is separate, but within the same paragraph.⋅⋅
⋅⋅⋅(This is contrary to the typical GFM line break behaviour, where trailing spaces are not required.)

* Unordered list can use asterisks
- Or minuses
+ Or pluses

***
## Blockquotes
```
> This is a very long line that will still be quoted properly when it wraps. Oh boy let's keep writing to make sure this is long enough to actually wrap for everyone. Oh, you can *put* **Markdown** into a blockquote. 

```
> This is a very long line that will still be quoted properly when it wraps. Oh boy let's keep writing to make sure this is long enough to actually wrap for everyone. Oh, you can *put* **Markdown** into a blockquote. 

***
## Language
```javascript
var s = "JavaScript syntax highlighting";
alert(s);
```
 
```python
s = "Python syntax highlighting"
print s
```
 
```
No language indicated, so no syntax highlighting. 
But let's throw in a <b>tag</b>.
```

***
## Rows and Columns
Colons can be used to align columns.

| Tables        | Are           | Cool  |
| ------------- | :-----------: | ----: |
| col 3 is      | right-aligned | $1600 |
| col 2 is      | centered      | $12   |
| zebra stripes | are neat      | $1    |

There must be at least 3 dashes separating each header cell.
The outer pipes (|) are optional, and you don't need to make the 
raw Markdown line up prettily. You can also use inline Markdown.

| Markdown | Less      | Pretty     |
| -------- | --------- | ---------- |
| *Still*  | `renders` | **nicely** |
| 1        | 2         | 3          |

***
## 강조
*single asterisks*
_single underscores_
**double asterisks**
__double underscores__
++underline++
~~cancelline~~

***
### 순서없는 목록(글머리 기호)
* Class
    * Method
        * *Return* 
    * Property
        * *Read / Write* 

***
## Youtube Player
```
[![Youtube Player](http://img.youtube.com/vi/유튜브 영상 키/0.jpg)](http://www.youtube.com/watch?v=유튜브 영상 키)
```
[![Youtube Player](http://img.youtube.com/vi/kJQP7kiw5Fk/0.jpg)](http://www.youtube.com/watch?v=kJQP7kiw5Fk)