import streamlit as st
import pandas as pd
from datetime import date, datetime
from PIL import Image
from to_excel import inputs_to_excel

# https://share.streamlit.io/streamlit/release-demos/0.84/0.84/streamlit_app.py?page=headliner
# https://docs.streamlit.io/library/advanced-features/session-state#initialization

def main():
    logo = Image.open('freedom_logo.png')

    st.set_page_config(
        page_title='Freedom PM & Sales CMA Form',
        page_icon=logo,
        layout='wide'
    )

    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        st.write('')
    with col2:
        st.image(logo)
    with col3:
        st.write('')

    disclaimer_container = st.container()
    password_container = st.container()
    description_container = st.container()
    instruction_container = st.container()
    intro_form_container = st.container()
    property_container = st.container()
    common_container = st.container()
    other_container = st.container()

    with disclaimer_container:
        with st.expander('DISCLOSURES'):
            st.markdown(
                '''
            *These estimates are not guaranteed and may not include escrows. Escrow balances are reimbursed by the existing lender. Taxes, rents & association dues are pro-rated at settlement. Under Virginia Law, the seller's proceeds may not be available for up to 2 business days following recording of the deed. Seller acknowledges receipt of this statement.*
            '''
            )

    with password_container:
        password_guess = st.text_input('Enter a password to gain access to this app', key='password_guess')
        if password_guess != st.secrets['password']:
            st.stop()

    with description_container:
        with st.expander('App Description'):
            st.markdown(
                '''
                ##### Comparative Market Analysis (CMA) Data Inputs
                This application is used to capture pertinent data points related to a CMA, which will then:
                - Calculate the Estimated Total Net Proceeds
                - Produce the Estimated Total Net Proceeds form in Excel
                ''')

    with instruction_container:
        with st.expander('Instructions'):
            st.markdown(
                '''
                - Open the Common Data Form container to check if pre-set values load
                - If preset values do not load, refresh the app, enter app password and check the Common Data Form
                - Perform this process until preset values appear
                - Enter known data into applicable fields
                    - For data that is a percentage, enter the number as the percentage you want
                    - For example if the Listing Company's Compensation is 2.25%, enter 2.25 into the number input field
                - After data for each input container is entered, press the container's "Submit" button to load the data into the app
                - Once all data is entered into all input containers, press "Download Net Proceeds Form"
                - A MS Excel workbook will be generated and placed into your downloads folder
                - The print area of the Excel file has already been set to allow for easy printing to PDF or paper
                '''
            )

    if 'preparer' not in st.session_state:
        st.session_state['preparer'] = ''
        st.session_state['prep_date'] = date.today()

        st.session_state['seller_name'] = ''
        st.session_state['seller_address'] = ''
        st.session_state['rec_list_price'] = 0
        st.session_state['estimated_payoff_first_trust'] = 0
        st.session_state['estimated_payoff_second_trust'] = 0
        st.session_state['annual_tax_amt'] = 0
        st.session_state['prorated_tax_amt'] = 0.0
        st.session_state['annual_hoa_condo_amt'] = 0
        st.session_state['prorated_hoa_condo_amt'] = 0
        st.session_state['property_subtotal'] = 0.0

        st.session_state['down_payment_pct'] = 0.0

        st.session_state['update_listing_company_pct'] = 2.5
        st.session_state['listing_company_pct'] = 0.025
        st.session_state['update_selling_company_pct'] = 2.5
        st.session_state['selling_company_pct'] = 0.025
        st.session_state['update_processing_fee'] = 0
        st.session_state['processing_fee'] = 0
        st.session_state['update_settlement_fee'] = 450
        st.session_state['settlement_fee'] = 0
        st.session_state['update_deed_prep_fee'] = 150
        st.session_state['deed_prep_fee'] = 0
        st.session_state['update_lien_release_fee'] = 100
        st.session_state['lien_release_fee'] = 0
        st.session_state['update_lien_trust_qty'] = 1
        st.session_state['lien_trust_qty'] = 0
        st.session_state['update_recording_release_fee'] = 38
        st.session_state['recording_release_fee'] = 0
        st.session_state['update_release_qty'] = 1
        st.session_state['release_qty'] = 0
        st.session_state['update_grantors_tax_pct'] = 0.1
        st.session_state['grantors_tax_pct'] = 0.001
        st.session_state['update_congestion_tax_pct'] = 0.2
        st.session_state['congestion_tax_pct'] = 0.002
        st.session_state['update_pest_inspection_fee'] = 50
        st.session_state['pest_inspection_fee'] = 0
        st.session_state['update_poa_condo_disclosure_fee'] = 350
        st.session_state['poa_condo_disclosure_fee'] = 0
        st.session_state['common_subtotal'] = 0.0

        st.session_state['closing_subsidy_radio'] = 'Percent of Recommended List Price (%)'
        st.session_state['update_closing_subsidy_pct'] = 0.0
        st.session_state['closing_subsidy_flat_amt'] = 0
        st.session_state['closing_subsidy_pct'] = 0.0
        st.session_state['closing_subsidy_amt'] = 0
        st.session_state['other_fee_name'] = ''
        st.session_state['other_fee_amt'] = 0
        st.session_state['other_subtotal'] = 0.0
        st.session_state['net_proceeds'] = 0.0

    def update_intro_info_form():
        st.session_state.preparer = st.session_state.preparer
        st.session_state.prep_date = st.session_state.prep_date

    def update_property_info_form():
        st.session_state.seller_name = st.session_state.seller_name
        st.session_state.seller_address = st.session_state.seller_address
        st.session_state.rec_list_price = st.session_state.rec_list_price
        st.session_state.estimated_payoff_first_trust = st.session_state.estimated_payoff_first_trust
        st.session_state.estimated_payoff_second_trust = st.session_state.estimated_payoff_second_trust
        st.session_state.prorated_tax_amt = st.session_state.annual_tax_amt / 12 * 3
        st.session_state.prorated_hoa_condo_amt = st.session_state.annual_hoa_condo_amt / 12 * 3


    def update_common_info_form():
        st.session_state.listing_company_pct = st.session_state.update_listing_company_pct / 100
        st.session_state.selling_company_pct = st.session_state.update_selling_company_pct / 100
        st.session_state.processing_fee = st.session_state.update_processing_fee
        st.session_state.settlement_fee = st.session_state.update_settlement_fee
        st.session_state.deed_prep_fee = st.session_state.update_deed_prep_fee
        st.session_state.lien_release_fee = st.session_state.update_lien_release_fee
        st.session_state.lien_trust_qty = st.session_state.update_lien_trust_qty
        st.session_state.recording_release_fee = st.session_state.update_recording_release_fee
        st.session_state.release_qty = st.session_state.update_release_qty
        st.session_state.grantors_tax_pct = st.session_state.update_grantors_tax_pct / 100
        st.session_state.congestion_tax_pct = st.session_state.update_congestion_tax_pct / 100
        st.session_state.pest_inspection_fee = st.session_state.update_pest_inspection_fee
        st.session_state.poa_condo_disclosure_fee = st.session_state.update_poa_condo_disclosure_fee

    def update_other_info_form():
        st.session_state.closing_subsidy_pct = st.session_state.update_closing_subsidy_pct / 100
        if st.session_state.closing_subsidy_radio == 'Percent of Offer Amt (%)':
            st.session_state.closing_subsidy_amt = st.session_state.closing_subsidy_pct * st.session_state.rec_list_price
        else:
            st.session_state.closing_subsidy_amt = st.session_state.closing_subsidy_flat_amt
        st.session_state.other_fee_amt = st.session_state.other_fee_amt

    def net_proceeds():
        st.session_state.property_subtotal = (
                st.session_state.estimated_payoff_first_trust +
                st.session_state.estimated_payoff_second_trust +
                st.session_state.closing_subsidy_amt +
                st.session_state.prorated_tax_amt +
                st.session_state.prorated_hoa_condo_amt
        )
        st.session_state.common_subtotal = (
            (st.session_state.listing_company_pct * st.session_state.rec_list_price) +
            (st.session_state.selling_company_pct * st.session_state.rec_list_price) +
            (st.session_state.processing_fee) +
            (st.session_state.deed_prep_fee) +
            (st.session_state.lien_release_fee * st.session_state.lien_trust_qty) +
            (st.session_state.recording_release_fee * st.session_state.release_qty) +
            (st.session_state.grantors_tax_pct * st.session_state.rec_list_price) +
            (st.session_state.congestion_tax_pct * st.session_state.rec_list_price) +
            (st.session_state.pest_inspection_fee) +
            (st.session_state.poa_condo_disclosure_fee)
        )
        st.session_state.other_subtotal = (
            st.session_state.other_fee_amt
        )
        st.session_state.net_proceeds = (
                st.session_state.rec_list_price -
                (st.session_state.property_subtotal +
                 st.session_state.common_subtotal +
                 st.session_state.other_subtotal)
        )
        return st.session_state.net_proceeds

    with intro_form_container:
        with st.expander('Introduction Data Form'):
            with st.form(key='intro_info_form'):
                st.markdown('##### **Form Preparation Data**')
                intro_col1, intro_col2 = st.columns(2)
                with intro_col1:
                    st.text_input('Enter the preparing agent\'s name', key='preparer')
                with intro_col2:
                    st.date_input('Enter preparation date of the form', key='prep_date')
                intro_info_submit = st.form_submit_button('Submit Information', on_click=update_intro_info_form)

    with property_container:
        with st.expander('Property Data Form'):
            with st.form(key='property_info_form'):
                st.markdown('##### **Property-specific Data**')
                property_col1, property_col2 = st.columns(2)
                with property_col1:
                    st.text_input('Enter seller\'s name(s)', key='seller_name')
                    st.text_input("Enter seller's address", key='seller_address')
                    st.number_input("Recommended Listing Price ($)", 0, 1500000, step=1000, key='rec_list_price')
                with property_col2:
                    st.number_input("Estimated Payoff - First Trust ($)", 0, 1000000, step=1000, key='estimated_payoff_first_trust')
                    st.number_input("Estimated Payoff - Second Trust ($)", 0, 1000000, step=1000, key='estimated_payoff_second_trust')
                    st.number_input("Annual Tax Amount ($)", 0, 25000, step=1, key='annual_tax_amt')
                    st.number_input('Annual HOA / Condo Amount ($)', 0, 10000, step=1, key='annual_hoa_condo_amt')
                property_info_submit = st.form_submit_button('Submit Property Information', on_click=update_property_info_form)

    with common_container:
        with st.expander('Common Data Form'):
            with st.form(key='common_info_form'):
                st.markdown('##### **Information Common to All Considered Transactions**')
                brokerage_col, closing_cost_col, misc_col = st.columns(3)
                with brokerage_col:
                    st.markdown('###### **Brokerage Cost Data**')
                    st.number_input("Listing Company's Compensation (%)", 0.0, 6.0, step=0.01, format='%.2f', key='update_listing_company_pct')
                    st.number_input("Selling Company's Compensation (%)", 0.0, 6.0, step=0.01, format='%.2f', key='update_selling_company_pct')
                    st.number_input('Processing Fee Amount ($)', 0, 20000, step=1, key='update_processing_fee')
                with closing_cost_col:
                    st.markdown('###### **Closing Cost Data**')
                    st.number_input('Settlement Fee Amount ($)', 0, 1000, step=1, key='update_settlement_fee')
                    st.number_input('Deed Preparation Fee Amount ($)', 0, 1000, step=1, key='update_deed_prep_fee')
                    st.number_input('Release of Liens / Trusts Fee Amount ($)', 0, 1000, step=1, key='update_lien_release_fee')
                    st.number_input('Number of Liens / Trusts', 0, 10, step=1, key='update_lien_trust_qty')
                with misc_col:
                    st.markdown('###### **Miscellaneous Cost Data**')
                    st.number_input('Recording Release(s) Fee Amount ($)', 0, 250, step=1, key='update_recording_release_fee')
                    st.number_input('Number of Releases', 0, 10, step=1, key='release_qty')
                    st.number_input("Grantor's Tax (%)", 0.0, 1.0, step=0.01, format='%.2f', key='update_grantors_tax_pct')
                    st.number_input("Congestion Relief Tax (%)", 0.0, 1.0, step=0.01, format='%.2f', key='update_congestion_tax_pct')
                    st.number_input("Pest Inspection Fee Amount ($)", 0, 100, step=1, key='update_pest_inspection_fee')
                    st.number_input("POA / Condo Disclosure Fee Amount ($)", 0, 500, step=1, key='update_poa_condo_disclosure_fee')
                common_info_submit = st.form_submit_button('Submit Common Information', on_click=update_common_info_form)

    with other_container:
        with st.expander('Other Data Form'):
            with st.form(key='other_info_form'):
                st.markdown('###### Closing Subsidy Data')
                other_col1, other_col2 = st.columns(2)
                with other_col1:
                    st.radio('Closing Cost Subsidy Radio', ['Percent of Recommended List Price (%)', 'Flat $ Amount'], key='closing_subsidy_radio')
                with other_col2:
                    st.number_input('Closing Cost Subsidy of (%):', 0.0, 100.0, step=0.01, key='update_closing_subsidy_pct')
                    st.number_input('Closing Cost Subsidy of ($):', 0, 100000, step=50, key='closing_subsidy_flat_amt')
                st.write('---')
                st.markdown('###### Other Data to be Included')
                other_col3, other_col4 = st.columns(2)
                with other_col3:
                    st.text_input('Enter name of another fee, if applicable', key='other_fee_name')
                with other_col4:
                    st.number_input('Enter the amount for the \'Other\' fee, if applicable', 0, 100000, step=1000, key='other_fee_amt')
                other_info_submit = st.form_submit_button('Submit Other Information', on_click=update_other_info_form)

    # st.write(st.session_state)

    proceeds_form = inputs_to_excel(
        agent=st.session_state.preparer,
        date=st.session_state.prep_date,
        seller=st.session_state.seller_name,
        address=st.session_state.seller_address,
        first_trust=st.session_state.estimated_payoff_first_trust,
        second_trust=st.session_state.estimated_payoff_second_trust,
        annual_taxes=st.session_state.annual_tax_amt,
        prorated_taxes=st.session_state.prorated_tax_amt,
        annual_hoa_condo_amt=st.session_state.annual_hoa_condo_amt,
        prorated_annual_hoa_condo_amt=st.session_state.prorated_hoa_condo_amt,
        list_price=st.session_state.rec_list_price,
        down_payment_pct=st.session_state.down_payment_pct,
        closing_subsidy_amt=st.session_state.closing_subsidy_amt,
        listing_company_pct=st.session_state.listing_company_pct,
        selling_company_pct=st.session_state.selling_company_pct,
        processing_fee=st.session_state.processing_fee,
        settlement_fee=st.session_state.settlement_fee,
        deed_preparation_fee=st.session_state.deed_prep_fee,
        lien_release_fee=st.session_state.lien_release_fee,
        lien_trust_qty=st.session_state.lien_trust_qty,
        recording_release_fee=st.session_state.recording_release_fee,
        recording_release_qty=st.session_state.release_qty,
        grantors_tax_pct=st.session_state.grantors_tax_pct,
        congestion_tax_pct=st.session_state.congestion_tax_pct,
        pest_inspection_fee=st.session_state.pest_inspection_fee,
        poa_condo_disclosure_fee=st.session_state.poa_condo_disclosure_fee,
        other_fee_name=st.session_state.other_fee_name,
        other_fee_amt=st.session_state.other_fee_amt
    )

    # st.write(f'Estimated Total Net Proceeds: ${net_proceeds()}', )

    st.download_button(
        label='Download Net Proceeds Form',
        data=proceeds_form,
        mime='xlsx',
        file_name=f"net_proceeds_form_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )







if __name__ == '__main__':
    main()