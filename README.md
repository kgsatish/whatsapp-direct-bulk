# Whatsapp-Direct-Bulk-Sender

Whatsapp Direct Bulk Sender sends whatsapp messages, including attachments, to specified mobile numbers, without having to add them as contacts on the phone.

Refer to requirements.txt for the dependencies.
 
## Installation

Install the dependencies before running the python script. Use python 3.10 and above.

```bash
pip install -r requirments.txt
```
After that, clone the git repository, go to the root folder `whatsapp-direct-bulk` and run the python script as given in 'Usage' section.


## Usage
Type the following command at the terminal to get help.

```python
python app.py -h
```
Example: (use test.png and test.csv and replace phone numbers in 'Contact' column for before testing)
```python
python app.py --img "C:\software\python\whatsapp-bulk-direct\test.png test.csv
```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change. Reach me on satishkg@yahoo.com for any queries

## Roadmap
Currently, works for images. Video, audio, document attachments yet to be implemented. Also, tested only against Indian mobile numbers. International numbers yet to be tested. You are welcome to chip in if interested.

## Authors and acknowledgment
During development, I referred to [Anurag's project](https://github.com/anirudhbagri/whatsapp-bulk-messenger) and made further improvements

## License
[MIT](https://choosealicense.com/licenses/mit/)